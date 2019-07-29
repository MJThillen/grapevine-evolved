VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlayers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Information"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   9060
   Icon            =   "frmPlayers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9060
   Begin VB.CommandButton cmdChange 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8190
      Picture         =   "frmPlayers.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CheckBox chkHideInactive 
      Caption         =   "Hide Inacti&ve Players"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "A&ctive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox txtPoints 
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   960
      Width           =   1230
   End
   Begin VB.ComboBox cboPosition 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPlayers.frx":0B14
      Left            =   3600
      List            =   "frmPlayers.frx":0B24
      TabIndex        =   13
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdFindShow 
      Caption         =   "&Find Characters..."
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   3840
      Width           =   2055
   End
   Begin VB.ListBox lstCharacters 
      Height          =   1575
      IntegralHeight  =   0   'False
      ItemData        =   "frmPlayers.frx":0B4E
      Left            =   6720
      List            =   "frmPlayers.frx":0B50
      TabIndex        =   24
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtNotes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   2880
      Width           =   5175
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   2400
      Width           =   5175
   End
   Begin VB.TextBox txtEMail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   1920
      Width           =   5175
   End
   Begin VB.CommandButton cmdAddPlayer 
      Caption         =   "&Add New Player"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdDeletePlayer 
      Caption         =   "&Delete Player"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   240
      Width           =   4575
   End
   Begin VB.ListBox lstPlayers 
      Height          =   4065
      IntegralHeight  =   0   'False
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin MSComCtl2.UpDown updPoints 
      Height          =   375
      Left            =   8190
      TabIndex        =   11
      Top             =   960
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   661
      _Version        =   393216
      OrigLeft        =   6015
      OrigTop         =   1815
      OrigRight       =   6600
      OrigBottom      =   2130
      Max             =   32767
      Min             =   -32768
      Orientation     =   1
      Enabled         =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8280
      Picture         =   "frmPlayers.frx":0B52
      Top             =   307
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Pla&yer Points Unspent/Earned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   5400
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   12
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Player &ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   1035
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "No&tes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   21
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Add&ress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   19
      Top             =   2955
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "P&hone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   17
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   15
      Top             =   1995
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "P&layers"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Player As PlayerClass   'a reference to the currently selected player object

Private FindShow As Boolean     'whether to Find characters or Show a selected character
Private Populating As Boolean   'whether the player points field is populating
Private Const fsFind = False
Private Const fsShow = True
Private ShiftDown As Boolean

Private Sub ReFillPlayerList()
'
' Name:         ReFillPlayerList
' Description:  Refresh the list of player names from the player list.
'

    Dim StoreCursor As Integer
    
    StoreCursor = lstPlayers.ListIndex
    
    lstPlayers.Clear
    PlayerList.First
    Do Until PlayerList.Off
        If chkHideInactive.Value = vbChecked Then
            If PlayerList.Item.Active Then
                lstPlayers.AddItem PlayerList.Item.Name
            End If
        Else
            lstPlayers.AddItem PlayerList.Item.Name
        End If
        PlayerList.MoveNext
    Loop

    lblLabels(0) = CStr(lstPlayers.ListCount) & _
                   IIf(chkHideInactive.Value = vbChecked, " Active", "") & _
                   " P&layers"

    If StoreCursor >= lstPlayers.ListCount Then StoreCursor = lstPlayers.ListCount - 1
    lstPlayers.ListIndex = StoreCursor

End Sub

Private Sub SetDataChanged()
'
' Name:         SetDataChanged
' Description:  Tell the game its data has changed and update the character's
'               Last Modified date.
'
    If Not Populating Then
        Game.DataChanged = True
        Player.LastModified = Now
    End If
    
End Sub

Private Sub chkActive_Click()
'
' Name:         chkActive_Click
' Description:  Change the player's Active status.
'

    If lstPlayers.ListIndex <> -1 Then
        Player.Active = (chkActive.Value = vbChecked)
        mdiMain.AnnounceChanges Me, atPlayers
        SetDataChanged
        If Not Player.Active And chkHideInactive.Value = vbChecked Then
            Dim Store As Integer
            Store = lstPlayers.ListIndex
            lstPlayers.RemoveItem Store
            lblLabels(0) = CStr(lstPlayers.ListCount) & " Active P&layers"
            If Store >= lstPlayers.ListCount Then Store = lstPlayers.ListCount - 1
            lstPlayers.ListIndex = Store
        End If
    End If

End Sub

Private Sub chkHideInactive_Click()
'
' Name:         chkHideInactive_Click
' Description:  Refresh the player list, save the setting.
'

    If Not Populating Then
        SaveSetting App.Title, "Settings", "HideInactive", (chkHideInactive.Value = vbChecked)
        ReFillPlayerList
    End If

End Sub

Private Sub cmdAddPlayer_Click()
'
' Name:         cmdAddPlayer_Click
' Description:  Add a new player to the game.
'

    Dim NewName As String
    
    NewName = InputBox("Enter a name for the player.", "Add New Player")
    NewName = Trim(NewName)
    
    If NewName <> "" Then
    
        PlayerList.MoveTo NewName
        If PlayerList.Off Then
    
            Set Player = New PlayerClass
            Player.Name = NewName
            PlayerList.Append Player
            ReFillPlayerList
            lstPlayers.ListIndex = lstPlayers.NewIndex
            Call lstPlayers_Click
            txtID.SetFocus
        
            mdiMain.AnnounceChanges Me, atPlayers
            Game.DataChanged = True
        Else
            MsgBox "The name """ & NewName & """ is already in use.  Please " & _
                    "enter a different name.", vbExclamation + vbOKOnly, "Duplicate Name"
        End If
        
    End If

End Sub

Private Sub cmdChange_Click()
'
' Name:         cmdChange_Click
' Description:  Add a new entry to the player point history, and adjust
'               the PP display accordingly.
'

    If lstPlayers.ListIndex <> -1 Then
    
        frmHistoryEntry.MakeEntry Player.Experience, True, "Change Points (Add History Entry)"
        If Not frmHistoryEntry.Canceled Then
            SetDataChanged
            Populating = True
            txtPoints = " " & CStr(Player.Experience.Unspent) & _
                    " / " & CStr(Player.Experience.Earned)
            Populating = False
        End If

    End If

End Sub

Private Sub cmdDeletePlayer_Click()
'
' Name:         cmdDeletePlayer_Click
' Description:  Delete a player form the game.
'

    If lstPlayers.ListIndex <> -1 Then
        
        PlayerList.MoveTo lstPlayers.Text
        If Not PlayerList.Off Then
        
            Dim Answer As Boolean
            
            Answer = ShiftDown
            If Not Answer Then Answer = (MsgBox("Are you sure you want to delete the entry for " & _
                    lstPlayers.Text & "?", vbYesNo, "Delete Player") = vbYes)
            If Answer Then
                
                PlayerList.Remove
                mdiMain.AnnounceChanges Me, atPlayers
                Game.DataChanged = True
                ReFillPlayerList
                Call lstPlayers_Click
            
            End If
            
        Else
            MsgBox "PlayerList.MoveTo " & lstPlayers.Text & " failed!"
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

Private Sub cmdFindShow_Click()
'
' Name:         cmdFindShow_Click
' Description:  Clicked the first time, find all characters this player plays.
'               Clicked the second time, load the selected character sheet.
'

    If lstPlayers.ListIndex <> -1 Then
        
        If FindShow = fsFind Then
        
            CharacterList.First
            Do Until CharacterList.Off
                If CharacterList.Item.Player = Player.Name Then _
                    lstCharacters.AddItem CharacterList.Item.Name
                CharacterList.MoveNext
            Loop
            
            If lstCharacters.ListCount = 0 Then lstCharacters.AddItem "(none)"

            FindShow = fsShow
            cmdFindShow.Caption = "S&how Character"
        
        Else
        
            If lstCharacters.ListIndex <> -1 Then
                If lstCharacters.Text <> "(none)" Then _
                    mdiMain.ShowCharacterSheet lstCharacters.Text
            End If
        End If

    End If

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Refresh the current player's point total.
'

    If lstPlayers.ListIndex <> -1 Then
        Populating = True
        txtPoints = " " & CStr(Player.Experience.Unspent) & _
                " / " & CStr(Player.Experience.Earned)
        Populating = False
    End If
    txtPoints.Locked = Game.EnforceHistory
    cmdChange.Visible = Game.EnforceHistory
    updPoints.Visible = Not Game.EnforceHistory

End Sub

Private Sub Form_Deactivate()
'
' Name:         Form_Deactivate
' Description:  When the window loses focus, reset the find/show characters list
'               and validate any remaining entries.
'
    
    Me.ValidateControls
    FindShow = fsFind
    cmdFindShow.Caption = "&Find Characters..."
    lstCharacters.Clear

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
' Name:         Form_KeyDown
' Description:  Record the state of the Shift key for deletions.
'

    If KeyCode = vbKeyShift Then ShiftDown = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'
' Name:         Form_KeyDown
' Description:  Record the state of the Shift key for deletions.
'

    If KeyCode = vbKeyShift Then ShiftDown = False

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the window's controls and data.
'

    Dim DataState As Boolean
    
    Populating = True
    DataState = Game.DataChanged
    chkHideInactive.Value = _
        IIf(GetSetting(App.Title, "Settings", "HideInactive", False), vbChecked, vbUnchecked)
    ReFillPlayerList
    If lstPlayers.ListCount > 0 Then lstPlayers.ListIndex = 0
    Game.DataChanged = DataState
    Populating = False
    mdiMain.OrientForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Ensure the data from the last control is valid.
'

    If Me.ActiveControl Is txtName Then
        Dim BadName As Boolean
        Call txtName_Validate(BadName)
        Cancel = BadName
    Else
        ValidateControls
    End If
    
End Sub

Private Sub lstCharacters_DblClick()
'
' Name:         lstCharacters_DblClick
' Description:  Load the clicked character sheet.
'

    Call cmdFindShow_Click

End Sub

Private Sub lstPlayers_Click()
'
' Name:         lstPlayers_Click
' Description:  Select a player, populating all fields with its data.
'

    Populating = True
    If lstPlayers.ListIndex <> -1 Then
        
        PlayerList.MoveTo lstPlayers.Text
        If Not PlayerList.Off Then
        
            Set Player = PlayerList.Item
            txtName.Text = Player.Name
            txtID.Text = Player.ID
            txtPoints.Text = " " & CStr(Player.Experience.Unspent) & _
                    " / " & CStr(Player.Experience.Earned)
            cboPosition.Text = Player.Position
            chkActive.Value = IIf(Player.Active, vbChecked, vbUnchecked)
            txtEMail.Text = Player.EMail
            txtPhone.Text = Player.Phone
            txtAddress.Text = Player.Address
            txtNotes.Text = Player.Notes
        
        Else
            MsgBox "PlayerList.MoveTo " & lstPlayers.Text & " failed!"
        End If
            
    Else
    
        txtName.Text = ""
        txtEMail.Text = ""
        txtID.Text = ""
        txtPoints.Text = ""
        cboPosition.Text = ""
        txtPhone.Text = ""
        txtAddress.Text = ""
        txtNotes.Text = ""
        
    End If
    
    FindShow = fsFind
    cmdFindShow.Caption = "&Find Characters..."
    lstCharacters.Clear
    Populating = False
    
End Sub

Private Sub lstPlayers_DblClick()
'
' Name:         lstPlayers_DblClick
' Description:  Select the text of the player's name.
'

    chkActive.Value = IIf(chkActive.Value = vbChecked, vbUnchecked, vbChecked)
    chkActive.SetFocus

End Sub

Private Sub lstPlayers_KeyUp(KeyCode As Integer, Shift As Integer)
'
' Name:         lstPlayers_KeyUp
' Description:  Process a delete key press.
'
    If KeyCode = vbKeyDelete Then Call cmdDeletePlayer_Click

End Sub

Private Sub txtEMail_GotFocus()
'
' Name:         txtEMail_GotFocus
' Description:  Select the text.
'

    SelectText txtEMail
    
End Sub

Private Sub txtEMail_KeyPress(KeyAscii As Integer)
'
' Name:         txtEMail_KeyPress
' Description:  Move to the next field when return is pressed.
'
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtEMail_Validate(False)
        txtPhone.SetFocus
    End If

End Sub

Private Sub txtAddress_GotFocus()
'
' Name:         txtAddress_GotFocus
' Description:  Select the text.
'

    SelectText txtAddress
    
End Sub

Private Sub txtName_GotFocus()
'
' Name:         txtName_GotFocus
' Description:  Select the text.
'

    SelectText txtName

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
'
' Name:         txtName_KeyPress
' Description:  Move to the next field when return is pressed.
'
    
    If KeyAscii = vbKeyReturn Then
        Dim Cancel As Boolean
        KeyAscii = 0
        Call txtName_Validate(Cancel)
        If Not Cancel Then txtID.SetFocus
    End If

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
'
' Name:         txtName_Validate
' Parameters:   Cancel      whether or not this is an invalid value
' Description:  If the value has changed, store it.
'

    If lstPlayers.ListIndex <> -1 Then
    
        txtName = Trim(txtName)
    
        If txtName <> Player.Name Then
        
            PlayerList.MoveTo txtName
            If PlayerList.Off Then
                Player.Name = txtName
                mdiMain.AnnounceChanges Me, atPlayers
                SetDataChanged
                ReFillPlayerList
            Else
                MsgBox "The name """ & txtName & """ is already in use.  Please " & _
                        "enter a different name.", vbExclamation + vbOKOnly, "Duplicate Name"
                txtName = Player.Name
                Cancel = True
            End If
            
        End If
        
    End If
    
End Sub

Private Sub txtID_GotFocus()
'
' Name:         txtID_GotFocus
' Description:  Select the text.
'

    SelectText txtID

End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
'
' Name:         txtID_KeyPress
' Description:  Move to the next field when return is pressed.
'
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtID_Validate(False)
        txtPoints.SetFocus
    End If

End Sub

Private Sub txtID_Validate(Cancel As Boolean)
'
' Name:         txtID_Validate
' Parameters:   Cancel      whether or not this is an invalid value
' Description:  If the value has changed, store it.
'

    If lstPlayers.ListIndex <> -1 Then
        If txtID <> Player.ID Then
            Player.ID = txtID
            SetDataChanged
        End If
    End If
    
End Sub

Private Sub cboPosition_GotFocus()
'
' Name:         cboPosition_GotFocus
' Description:  Select the text.
'

    cboPosition.SelStart = 0
    cboPosition.SelLength = Len(cboPosition.Text)

End Sub

Private Sub cboPosition_KeyPress(KeyAscii As Integer)
'
' Name:         cboPosition_KeyPress
' Description:  Move to the next field when return is pressed.
'
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call cboPosition_Validate(False)
        txtEMail.SetFocus
    End If

End Sub

Private Sub cboPosition_Validate(Cancel As Boolean)
'
' Name:         cboPosition_Validate
' Parameters:   Cancel      whether or not this is an invalid value
' Description:  If the value has changed, store it.  Adjust the Narrator list.
'

    If lstPlayers.ListIndex <> -1 Then
        If cboPosition <> Player.Position Then
            
            Player.Position = cboPosition
            mdiMain.AnnounceChanges Me, atPlayers
            SetDataChanged
        
        End If
    End If
    
End Sub

Private Sub txtEMail_Validate(Cancel As Boolean)
'
' Name:         txtEMail_Validate
' Parameters:   Cancel      whether or not this is an invalid value
' Description:  If the value has changed, store it.
'

    If lstPlayers.ListIndex <> -1 Then
        If Player.EMail <> txtEMail Then
            Player.EMail = txtEMail
            SetDataChanged
        End If
    End If

End Sub

Private Sub txtNotes_GotFocus()
'
' Name:         txtNotes_GotFocus
' Description:  Select the text.
'

    SelectText txtNotes
    
End Sub

Private Sub txtPhone_GotFocus()
'
' Name:         txtPhone_GotFocus
' Description:  Select the text.
'

    SelectText txtPhone
    
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
'
' Name:         txtPhone_KeyPress
' Description:  Move to the next field when return is pressed.
'

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtPhone_Validate(False)
        txtAddress.SetFocus
    End If

End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
'
' Name:         txtPhone_Validate
' Parameters:   Cancel      whether or not this is an invalid value
' Description:  If the value has changed, store it.
'

    If lstPlayers.ListIndex <> -1 Then
        If Player.Phone <> txtPhone Then
            Player.Phone = txtPhone
            SetDataChanged
        End If
    End If
        
End Sub

Private Sub txtAddress_Validate(Cancel As Boolean)
'
' Name:         txtAddress_Validate
' Parameters:   Cancel      whether or not this is an invalid value
' Description:  If the value has changed, store it.
'

    If lstPlayers.ListIndex <> -1 Then
        If Player.Address <> txtAddress Then
            Player.Address = txtAddress
            SetDataChanged
        End If
    End If
    
End Sub

Private Sub txtNotes_Validate(Cancel As Boolean)
'
' Name:         txtNotes_Validate
' Parameters:   Cancel      whether or not this is an invalid value
' Description:  If the value has changed, store it.
'

    If lstPlayers.ListIndex <> -1 Then
        If Player.Notes <> txtNotes Then
            Player.Notes = txtNotes
            SetDataChanged
        End If
    End If

End Sub

Private Sub txtPoints_Change()
'
' Name:         txtPoints_Change
' Description:  Ensure a valid value and save the change to the players's
'               Points.
'
    If lstPlayers.ListIndex <> -1 And Not Populating Then
        
        Dim Slash As Integer
        Dim Estr As String
        Dim Ustr As String
        
        Slash = InStr(txtPoints.Text, "/")
        
        If Slash > 0 Then
            
            Ustr = Trim(Left(txtPoints.Text, Slash - 1))
            Estr = Trim(Mid(txtPoints.Text, Slash + 1))
            
            If (IsNumeric(Ustr) Or Ustr = "") And _
               (IsNumeric(Estr) Or Estr = "") Then
                Player.Experience.Unspent = Val(Ustr)
                Player.Experience.Earned = Val(Estr)
                txtPoints.ForeColor = vbWindowText
                SetDataChanged
            Else
                txtPoints.ForeColor = vbHighlight
            End If
            
        Else
            txtPoints.Text = " " & CStr(Player.Experience.Unspent) & _
                    " / " & CStr(Player.Experience.Earned)
        End If
        
    End If

End Sub

Private Sub txtPoints_GotFocus()
'
' Name:         txtPoints_GotFocus
' Description:  Select the text.
'

    SelectText txtPoints

End Sub

Private Sub txtPoints_KeyPress(KeyAscii As Integer)
'
' Name:         txtPoints_KeyPress
' Description:  Move to the next field when return is pressed.
'

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cboPosition.SetFocus
    End If

End Sub

Private Sub updPoints_DownClick()
'
' Name:         updPoints_DownClick
' Description:  Update the label and store the new value.
'

    Dim EditBoth As Boolean
    Dim EditUnspent As Boolean
    Dim SaveStart As Integer
    Dim SaveLen As Integer
    
    If lstPlayers.ListIndex <> -1 Then
    
        Populating = True
        
        With txtPoints
        
            SaveStart = .SelStart
            SaveLen = .SelLength
            If SaveLen = Len(.Text) Then SaveLen = SaveLen + 2
            EditBoth = InStr(.SelText, "/") > 0 Or Not Me.ActiveControl Is txtPoints
            EditUnspent = SaveStart <= InStr(.Text, "/")

            With Player.Experience
                If EditBoth Or EditUnspent Then   'take from unspent
                    .Unspent = .Unspent - 1
                Else                               'take from earned
                    .Earned = .Earned - 1
                End If
                    
                txtPoints.Text = " " & CStr(.Unspent) & " / " & CStr(.Earned)
                        
            End With
            
            .SelStart = SaveStart
            .SelLength = SaveLen
            
        End With
            
        Populating = False
        
        SetDataChanged

    End If

End Sub

Private Sub updPoints_UpClick()
'
' Name:         updPoints_UpClick
' Description:  Update the label and store the new value.
'

    Dim EditBoth As Boolean
    Dim EditUnspent As Boolean
    Dim SaveStart As Integer
    Dim SaveLen As Integer
    
    If lstPlayers.ListIndex <> -1 Then
    
        Populating = True
        
        With txtPoints
        
            SaveStart = .SelStart
            SaveLen = .SelLength
            If SaveLen = Len(.Text) Then SaveLen = SaveLen + 2
            EditBoth = InStr(.SelText, "/") > 0 Or Not Me.ActiveControl Is txtPoints
            EditUnspent = SaveStart <= InStr(.Text, "/")

            With Player.Experience
                If EditBoth Then   'add to both
                    .Unspent = .Unspent + 1
                    .Earned = .Earned + 1
                ElseIf EditUnspent Then  'add to unspent
                    .Unspent = .Unspent + 1
                Else                'add to earned
                    .Earned = .Earned + 1
                End If
                    
                txtPoints.Text = " " & CStr(.Unspent) & " / " & CStr(.Earned)
                        
            End With
            
            .SelStart = SaveStart
            .SelLength = SaveLen
            
        End With
            
        Populating = False
        
        SetDataChanged
    
    End If
    
End Sub
