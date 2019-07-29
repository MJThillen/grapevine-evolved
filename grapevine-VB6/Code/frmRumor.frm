VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRumor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rumor"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmRumor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   9060
   Tag             =   "U"
   Begin VB.CheckBox chkDone 
      Alignment       =   1  'Right Justify
      Caption         =   "All D&one"
      Height          =   195
      Left            =   6000
      TabIndex        =   25
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame fraRumor 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   8535
      Begin VB.CommandButton cmdShowRecipient 
         Caption         =   "S&how Recipient"
         Height          =   375
         Left            =   6840
         TabIndex        =   14
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtRumor 
         Height          =   2775
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   4935
      End
      Begin VB.ListBox lstCauses 
         Height          =   1815
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.ListBox lstRecipients 
         Height          =   2295
         IntegralHeight  =   0   'False
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddCause 
         Caption         =   "&Add"
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveCause 
         Caption         =   "Remo&ve"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblSubTitle 
         Caption         =   "01/01/01 Level 1 Street &Rumor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Affected &By:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblRecipients 
         Caption         =   "0 Active Re&cipients:"
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdEscClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame fraQuery 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton cmdDeleteTerm 
         Caption         =   "&Delete Term"
         Height          =   375
         Left            =   6840
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddTerm 
         Caption         =   "Add T&erm"
         Height          =   375
         Left            =   6840
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.ListBox lstTerms 
         Height          =   855
         IntegralHeight  =   0   'False
         Left            =   3600
         TabIndex        =   17
         Top             =   120
         Width           =   3135
      End
      Begin VB.OptionButton optAnyAll 
         Caption         =   "A&ny of these terms:"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optAnyAll 
         Caption         =   "All of these terms:"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "&Send this rumor to characters that match"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Frame fraLevels 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      Begin MSComctlLib.TabStrip tabLevels 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   661
         TabWidthStyle   =   2
         Style           =   1
         TabFixedWidth   =   1415
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   10
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &2"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &3"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &4"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &5"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &6"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &7"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &8"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level &9"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Level 1&0"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "(*) Designates a rumor that has yet to be written."
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblHighest 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   165
         Width           =   375
      End
      Begin VB.Label lblLabel 
         Caption         =   "is the highest level of X among active characters."
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Label lblMainTitle 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      TabIndex        =   22
      Top             =   270
      Width           =   4935
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7080
      TabIndex        =   23
      Top             =   270
      Width           =   1695
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   840
      Picture         =   "frmRumor.frx":058A
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmRumor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MULTITOP = 1920
Private Const OPT_ANY = 0
Private Const OPT_ALL = 1

Private Rumor As RumorClass               'the rumor object this window manipulates
Private SubRumor As RumorNode             'the subrumor object presently under scrutiny
Private ShortDate As String
Private Populating As Boolean

Public Sub ShowRumor(Rum As RumorClass)
'
' Name:         ShowRumor
' Parameters:   Rum         the RumorClass object to show
' Description:  Show the given RumorClass object so the user can work with it.
'               Ready the form for its arrival.
'

    Set Rumor = Rum
    Populating = True
    
    Me.Caption = Rumor.Name
    ShortDate = Format(Rumor.RumorDate, "Short Date")
    lblMainTitle.Caption = " " & Rumor.Title
    lblDate.Caption = ShortDate
    imgIcon.Picture = mdiMain.imlIcons.ListImages(Rumor.IconKey).Picture
    chkDone.Value = IIf(Rumor.Done, vbChecked, vbUnchecked)
    
    If Rumor.Count = 0 Then
        Rumor.Add 0, ""
    End If
    
    Rumor.First
    Set SubRumor = Rumor.SubRumor
    
    If Rumor.Category = rtInfluence Then
    
        Dim I As Integer
        
        fraLevels.Visible = True
        fraRumor.Top = MULTITOP
        lblLabel(3).Caption = "is the highest level of " & Rumor.MultiMatch & _
                              " among active characters."
        
        Rumor.First
        Do Until Rumor.Off
            I = Rumor.SubRumor.Level
            If I >= 1 And I <= 10 Then
                tabLevels.Tabs(I).Caption = tabLevels.Tabs(I).Tag & _
                                            IIf(Rumor.SubRumor.Rumor = "", "*", "")
            End If
            Rumor.MoveNext
        Loop
        
        Set tabLevels.SelectedItem = Nothing
        Set tabLevels.SelectedItem = tabLevels.Tabs(SubRumor.Level)
        Call tabLevels_Click
        
    Else
        
        fraQuery.Visible = True
        lblSubTitle.Caption = ShortDate & " " & Rumor.Title & " &Rumor"
        txtRumor.Text = SubRumor.Rumor
        If Rumor.Query Is Nothing Then
            Set Rumor.Query = New QueryClass
            Rumor.Query.Inventory = qiCharacters
        End If
        optAnyAll(IIf(Rumor.Query.MatchAll, OPT_ALL, OPT_ANY)).Value = True
        RefreshTerms
        RefreshLists
                
    End If
        
    mdiMain.OrientForm Me

    Populating = False
    
    Me.Show

End Sub

Public Sub ShowLevel(Level As Integer)
'
' Name:         ShowLevel
' Parameter:    Level       Rumor level to show
' Description:  Allow other forms to force this form to jump to another rumor level.
'

    If Rumor.Category = rtInfluence Then
        ValidateControls
        Set tabLevels.SelectedItem = tabLevels.Tabs(Level)
        Call tabLevels_Click
    End If

End Sub

Private Sub RefreshLists()
'
' Name:         RefreshLists
' Description:  Refresh the cause and recipient lists.
'

    Dim Results As LinkedList
    Dim Values As LinkedList
    Dim Query As QueryClass

    Screen.MousePointer = vbHourglass
    Populating = True
    
    lstRecipients.Clear
    
    SubRumor.Causes.PopulateList lstCauses, Rumor.RumorDate

    If Rumor.Category = rtInfluence Then
        Set Query = New QueryClass
        Query.Inventory = qiCharacters
        Query.AddClause Rumor.MultiKey, Rumor.MultiMatch, SubRumor.Level, qcContainsAtLeast, False
    Else
        Set Query = Rumor.Query
    End If

    Game.QueryEngine.MakeQuery Query
    With Game.QueryEngine.Results
        .First
        Do Until .Off
            If .Item.Status = "Active" Then lstRecipients.AddItem .Item.Name
            .MoveNext
        Loop
    End With
    
    lblRecipients.Caption = CStr(lstRecipients.ListCount) & " Active Re&cipients:"
    
    Set Query = Nothing
    
    Populating = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub RefreshTerms()
'
' Name:         RefreshTerms
' Description:  Refresh all the terms of the query.
'

    Dim Sel As Integer
    Dim OldCategory As RumorCategoryType
    
    Sel = lstTerms.ListIndex
    OldCategory = Rumor.Category
    
    lstTerms.Clear
    With Rumor.Query
        If .IsEmpty Then
            lstTerms.AddItem "All Characters"
            Rumor.Category = rtGeneral
        Else
            .First
            Select Case .Clause.Key
                Case qkGroup:       Rumor.Category = rtGroup
                Case qkSubgroup:    Rumor.Category = rtSubgroup
                Case qkRace:        Rumor.Category = rtRace
                Case qkName:        Rumor.Category = rtPersonal
                Case Else:          Rumor.Category = rtGeneral
            End Select
            Do Until .Off
                lstTerms.AddItem .ClauseDescNext
            Loop
        End If
    End With
    
    If Rumor.Category <> OldCategory Then
        Rumor.LastModified = Now
        imgIcon.Picture = mdiMain.imlIcons.ListImages(Rumor.IconKey).Picture
        mdiMain.AnnounceChanges Me, atRumors
    End If
    
    If Sel >= lstTerms.ListCount Then Sel = lstTerms.ListCount - 1
    lstTerms.ListIndex = Sel
    
End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnActionRumor
        .SelectSet(osRumors).Clear
        .SelectSet(osRumors).Add Rumor.Name
        .GameDate = Rumor.RumorDate
    End With

End Sub

Private Sub chkDone_Click()
'
' Name:         chkDone_Click
' Description:  Set the Done status of this rumor.
'

    If Rumor.Done <> (chkDone.Value = vbChecked) Then
        Rumor.Done = (chkDone.Value = vbChecked)
        Game.DataChanged = True
        Rumor.LastModified = Now
        mdiMain.AnnounceChanges Me, atRumors
    End If
    
End Sub

Private Sub cmdAddCause_Click()
'
' Name:         cmdAddCause_Click
' Description:  Add a cause link to the subrumor.
'

    frmSelectLink.SelectLink Rumor.RumorDate, False
    If Not frmSelectLink.ChoiceType = aprNone Then
    
        SubRumor.Causes.AddLink frmSelectLink.ChoiceType, frmSelectLink.When, _
                                frmSelectLink.Item, frmSelectLink.Subitem
        SubRumor.Causes.PopulateList lstCauses, Rumor.RumorDate
        Rumor.LastModified = Now
        Game.DataChanged = True
        
    End If

End Sub

Private Sub cmdAddTerm_Click()
'
' Name:         cmdAddTerm_Click
' Description:  Add a new term to the query.
'

    frmQueryTerm.AddQueryTerm Rumor.Query
    RefreshTerms
    RefreshLists
    Game.DataChanged = True
    Rumor.LastModified = Now
    
End Sub

Private Sub cmdDeleteTerm_Click()
'
' Name:         cmdDeleteTerm_Click
' Description:  Delete the selected term of the query.
'

    If lstTerms.ListIndex > -1 And Not Rumor.Query.IsEmpty Then
        Rumor.Query.Remove lstTerms.ListIndex
        RefreshTerms
        RefreshLists
        Game.DataChanged = True
        Rumor.LastModified = Now
    End If

End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'
    Unload Me

End Sub

Private Sub cmdRemoveCause_Click()
'
' Name:         cmdRemoveCause_Click
' Description:  Remove the selected cause from the cause list.
'

    If lstCauses.ListIndex > -1 Then
    
        SubRumor.Causes.MoveToPlace lstCauses.ItemData(lstCauses.ListIndex)
        SubRumor.Causes.RemoveLink
        SubRumor.Causes.PopulateList lstCauses, Rumor.RumorDate
        Rumor.LastModified = Now
        Game.DataChanged = True
        
    End If

End Sub

Private Sub cmdShowRecipient_Click()
'
' Name:         cmdShowRecipient_Click
' Description:  Show the character sheet of the selected recipient.
'

    If lstRecipients.ListIndex > -1 Then
        mdiMain.ShowCharacterSheet lstRecipients.Text
    End If

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Characters may have changed, so refresh the lists and
'               check the highest value of the Influence/Background,
'               if this is a multirumor.
'
    
    RefreshLists
    If Rumor.Category = rtInfluence Then
        
        Dim Query As QueryClass
        Dim Max As Double
        
        Screen.MousePointer = vbHourglass
        
        Game.QueryEngine.QueryList.MoveTo "Active Characters"
        If Game.QueryEngine.QueryList.Off Then
            Set Query = New QueryClass
            Query.Inventory = qiCharacters
            Query.AddClause qkPlayStatus, "Active", 0, qcEquals, False
        Else
            Set Query = Game.QueryEngine.QueryList.Item
        End If
        
        With Game.QueryEngine
            .GetStatistics stMaxima, Query, Rumor.MultiKey
            
            Max = 0
            On Error Resume Next
            Max = .NumberSet(Rumor.MultiMatch)
            On Error GoTo 0
        End With
        
        lblHighest.Caption = CStr(Max)
        
        Set Query = Nothing
        Screen.MousePointer = vbDefault
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Validate the controls before the form is unloaded.
'
    ValidateControls

End Sub

Private Sub lblDate_Click()
'
' Name:         lblMainTitle_Click
' Description:  Offer the option to reassign title/dates if clicking the date.
'
    Call lblMainTitle_Click
    
End Sub

Private Sub lblMainTitle_Click()
'
' Name:         lblMainTitle_Click
' Description:  Offer the option to reassign title/dates if clicking the title.
'

    Dim Choices As StringSet
    Dim NewDate As Date
    Dim NewName As String
    
    Set Choices = New StringSet
    
    Choices.Add "Change title"
    Choices.Add "Change date"
    
    Select Case mdiMain.CreatePopup(Choices, Me)
        Case "Change title"
            frmGetAPRInfo.GetNewRumorTitle Rumor.RumorDate
        Case "Change date"
            frmGetAPRInfo.GetNewRumorDate Rumor.Title
        Case Else
            Exit Sub
    End Select
    
    NewDate = frmGetAPRInfo.NewDate
    NewName = Trim(frmGetAPRInfo.NewItem)
    Unload frmGetAPRInfo
    
    If Not (NewName = "" Or (NewName = Rumor.Title And NewDate = Rumor.RumorDate)) Then
    
        Game.APREngine.MoveToPair RumorList, NewDate, NewName
        If RumorList.Off Then
            Game.APREngine.Reassign RumorList, Rumor.Title, NewName, Rumor.RumorDate, NewDate
            ShortDate = Format(Rumor.RumorDate, "Short Date")
            Me.Caption = Rumor.Name
            lblMainTitle.Caption = " " & Rumor.Title
            lblDate.Caption = ShortDate
            SubRumor.Causes.PopulateList lstCauses, Rumor.RumorDate
            If Rumor.Category = rtInfluence Then
                lblSubTitle.Caption = ShortDate & " Level " & CStr(SubRumor.Level) _
                                    & " " & Rumor.Title & " &Rumor"
            Else
                lblSubTitle.Caption = ShortDate & " " & Rumor.Title & " &Rumor"
            End If
            Rumor.LastModified = Now
            Game.DataChanged = True
            mdiMain.AnnounceChanges Me, atRumors
        Else
            MsgBox "A rumor already exists under that title for that date.", _
                   vbOKOnly + vbExclamation, "Reassign Rumor"
        End If
            
    End If
    
    Set Choices = Nothing

End Sub

Private Sub lstCauses_Click()
'
' Name:         lstCauses_Click
' Description:  Set the tooltip of this control to match the selection,
'               helping the user view captions that don't fit.
'

    lstCauses.ToolTipText = lstCauses.Text

End Sub

Private Sub lstCauses_DblClick()
'
' Name:         lstCauses_Click
' Description:  Show the window associated with the link double-clicked.
'

    If lstCauses.ListIndex > -1 Then
                
        SubRumor.Causes.MoveToPlace lstCauses.ItemData(lstCauses.ListIndex)
        
        If Not SubRumor.Causes.Off Then
            With SubRumor.Causes.Link
                mdiMain.ShowAPR .Target, .Item, .When, .Subitem
            End With
        End If
                
    End If

End Sub

Private Sub lstCauses_KeyPress(KeyAscii As Integer)
'
' Name:         lstCauses_KeyPress
' Description:  Translate an enter into a double-click.
'
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call lstCauses_DblClick
    End If

End Sub

Private Sub lstRecipients_DblClick()
'
' Name:         lstRecipients_DblClick
' Description:  Shortcut to cmdShowRecipient_Click
'
    Call cmdShowRecipient_Click

End Sub

Private Sub lstTerms_DblClick()
'
' Name:         lstTerms_DblClick
' Description:  Edit the selected query term, or add if there is none.
'
    If Rumor.Query.IsEmpty Then
        cmdAddTerm_Click
    Else
        Rumor.Query.MoveToClause lstTerms.ListIndex
        frmQueryTerm.EditQueryTerm Rumor.Query
        RefreshTerms
        RefreshLists
        Game.DataChanged = True
        Rumor.LastModified = Now
    End If

End Sub

Private Sub optAnyAll_Click(Index As Integer)
'
' Name:         optAnyAll_Click
' Description:  Set whether this query matches all or any of the terms
'

    If Not Populating Then
        Rumor.Query.MatchAll = (Index = OPT_ALL)
        RefreshTerms
        RefreshLists
        Game.DataChanged = True
        Rumor.LastModified = Now
    End If

End Sub

Private Sub tabLevels_Click()
'
' Name:         tabLevels_Click
' Description:  Show the rumor associated with this level.
'
    
    Dim Index As String
    
    Index = CStr(tabLevels.SelectedItem.Index)
    If Not tabLevels.Tag = Index Then
    
        tabLevels.Tag = Index
        Rumor.MoveTo tabLevels.SelectedItem.Index
        If Not Rumor.Off Then
            Set SubRumor = Rumor.SubRumor
            lblSubTitle.Caption = ShortDate & " Level " & Index & " " & Rumor.Title & " &Rumor"
            txtRumor.Text = SubRumor.Rumor
            RefreshLists
        Else
            MsgBox "Something just broke.  The level " & Index & " rumor was never created!"
        End If
    
    End If

End Sub

Private Sub txtRumor_Validate(Cancel As Boolean)
'
' Name:         txtRumor_Validate
' Description:  Store the text for this rumor.
'

    txtRumor.Text = TrimWhiteSpace(txtRumor.Text)
    If Not (SubRumor.Rumor = txtRumor.Text) Then
        
        Dim I As Integer
        Dim OldDone As Boolean
        
        SubRumor.Rumor = txtRumor.Text
        OldDone = Rumor.Done
        I = SubRumor.Level
        If I > 0 Then
            tabLevels.Tabs(I).Caption = tabLevels.Tabs(I).Tag & IIf(Rumor.SubRumor.Rumor = "", "*", "")
            Rumor.IfDoneSetDone CInt(lblHighest.Caption), I
        Else
            Rumor.IfDoneSetDone 0, 0
        End If
        chkDone.Value = IIf(Rumor.Done, vbChecked, vbUnchecked)
        If Not OldDone = Rumor.Done Then mdiMain.AnnounceChanges Me, atRumors
        Game.DataChanged = True
        Rumor.LastModified = Now
    
    End If

End Sub
