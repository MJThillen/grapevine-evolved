VERSION 5.00
Begin VB.Form frmAddNewCharacter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Character"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7920
   ControlBox      =   0   'False
   Icon            =   "frmAddNewCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame fraTraits 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtRandom 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   18
         Text            =   "7"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtRandom 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Text            =   "5"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtRandom 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         Text            =   "3"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtRandom 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   22
         Text            =   "5"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtRandom 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   27
         Text            =   "5"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtRandom 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   3960
         TabIndex        =   29
         Text            =   "5"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtRandom 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   6120
         TabIndex        =   24
         Text            =   "5"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblRandom 
         Caption         =   "Attribute &Traits"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   17
         Top             =   45
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         Caption         =   "Ab&ility Traits"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   21
         Top             =   45
         Width           =   1455
      End
      Begin VB.Label lblRandom 
         Caption         =   "&Negative Traits"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   26
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label lblRandom 
         Caption         =   "&Background Traits"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   28
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label lblRandom 
         Caption         =   "Fr&ee Traits"
         Height          =   255
         Index           =   4
         Left            =   6600
         TabIndex        =   23
         Top             =   45
         Width           =   975
      End
      Begin VB.Label lblRandom 
         Alignment       =   1  'Right Justify
         Caption         =   "Up to"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   25
         Top             =   405
         Width           =   855
      End
      Begin VB.Label lblRandom 
         Alignment       =   2  'Center
         Caption         =   "/         /"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   16
         Top             =   45
         Width           =   1335
      End
      Begin VB.Label lblRandom 
         Caption         =   "Grapevine cannot ensure that the traits randomly chosen will make any sense at all."
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   31
         Top             =   840
         Width           =   6375
      End
      Begin VB.Label lblRandom 
         Alignment       =   1  'Right Justify
         Caption         =   "Note:"
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
         Index           =   8
         Left            =   0
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkRandom 
      Caption         =   "&Generate random basic traits"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton cmdDemon 
      Caption         =   "&Demon"
      Height          =   855
      Left            =   5640
      Picture         =   "frmAddNewCharacter.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdHunter 
      Caption         =   "&Hunter"
      Height          =   855
      Left            =   6720
      Picture         =   "frmAddNewCharacter.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdKueiJin 
      Caption         =   "&Kuei-Jin"
      Height          =   855
      Left            =   3480
      Picture         =   "frmAddNewCharacter.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdMummy 
      Caption         =   "M&ummy"
      Height          =   855
      Left            =   4560
      Picture         =   "frmAddNewCharacter.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdVarious 
      Caption         =   "V&arious"
      Height          =   855
      Left            =   6720
      Picture         =   "frmAddNewCharacter.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdFera 
      Caption         =   "&Fera"
      Height          =   855
      Left            =   2400
      Picture         =   "frmAddNewCharacter.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdMage 
      Caption         =   "&Mage"
      Height          =   855
      Left            =   3480
      Picture         =   "frmAddNewCharacter.frx":3D86
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdWraith 
      Caption         =   "&W&raith"
      Height          =   855
      Left            =   4560
      Picture         =   "frmAddNewCharacter.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdChangeling 
      Caption         =   "&Changeling"
      Height          =   855
      Left            =   5640
      Picture         =   "frmAddNewCharacter.frx":4F1A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   555
      Left            =   6720
      TabIndex        =   32
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdMortal 
      Caption         =   "M&ortal"
      Height          =   855
      Left            =   1320
      Picture         =   "frmAddNewCharacter.frx":57E4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdWerewolf 
      Caption         =   "&Werewolf"
      Height          =   855
      Left            =   2400
      Picture         =   "frmAddNewCharacter.frx":60AE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdVampire 
      Caption         =   "&Vampire"
      Height          =   855
      Left            =   1320
      Picture         =   "frmAddNewCharacter.frx":6978
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
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
      Height          =   555
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label lblName 
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
      Left            =   360
      TabIndex        =   0
      Top             =   345
      Width           =   855
   End
End
Attribute VB_Name = "frmAddNewCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Name:         frmAddNewCharacter
' Description:  Form with which you choose a name and race for a new character.
'               Only shown from frmCharacters, and always shown modal.
'

Public Race As RaceType         'Public data with which frmCharacters learns the character race;
                                '   equals gvRaceNone if Cancel is chosen.
Public CharacterName As String  'Public data with which frmCharacters learns the character name
                                '   equals "" if Cancel is chosen.
Public RandomGen As Boolean     'Whether to generate random traits for a new character

Private Const SHORT_HEIGHT = 3570   'Height of the window when Generate Random Traits is not checked
Private Const TALL_HEIGHT = 4890    'Height of the window when Generate Random Traits is checked

Private Const tiPrimary = 0
Private Const tiSecondary = 1
Private Const tiTertiary = 2
Private Const tiNegatives = 3
Private Const tiAbilities = 4
Private Const tiBackgrounds = 5
Private Const tiFree = 6

Public Sub GetCharacter(Optional DefaultName As String = "")
'
' Name:         GetCharacter
' Description:  Clears the variables and shows the New Character form modal.
' Returns:      Indirectly via Race and CharacterName, a race and name for the character.
'
    
    If DefaultName = "" Then
        Me.Caption = "Add New Character"
        chkRandom.Visible = True
        chkRandom.Value = IIf(RandomGen, vbChecked, vbUnchecked)
        Me.Height = IIf(RandomGen, TALL_HEIGHT, SHORT_HEIGHT)
        fraTraits.Visible = RandomGen
    Else
        Me.Caption = "Convert Character"
        chkRandom.Visible = False
        Me.Height = SHORT_HEIGHT
    End If
    
    txtName = DefaultName
    Race = gvRaceNone
    CharacterName = DefaultName
    Me.Show vbModal, mdiMain
    
End Sub

Private Sub SetCharacter(CharRace As RaceType)
'
' Name:         SetCharacter
' Description:  Validates the information and places the chosen name and race in
'               variables where frmCharacters will find them.
' Arguments:    an integer indicating the race of the character.
' Returns:      the race and name of the character, in Public variables.
'

    txtName.Text = Trim(txtName.Text)

    If txtName.Text = "" Then
        MsgBox "You must provide a name for the character.", _
                vbOKOnly, "Add New Character"
        txtName.SetFocus
    Else
        CharacterList.MoveTo txtName.Text
        If CharacterList.Off Then
            
            CharacterName = txtName.Text
            Race = CharRace
            If RandomGen Then
                Game.RandomTraits = txtRandom(tiPrimary).Text & "," & txtRandom(tiSecondary).Text & _
                              "," & txtRandom(tiTertiary).Text & "," & txtRandom(tiNegatives).Text & _
                              "," & txtRandom(tiAbilities).Text & "," & txtRandom(tiBackgrounds).Text & _
                              "," & txtRandom(tiFree).Text
            End If
            SaveSetting App.Title, "Settings", "Random Traits", RandomGen
            Me.Hide
        
        Else
            MsgBox "The name """ & Trim(txtName) & """ is already" & _
                   " in use.  Please choose another name.", _
                    vbOKOnly, "Duplicate Character Name"
            txtName.SetFocus
        End If
    End If

End Sub

Private Sub chkRandom_Click()
'
' Name:         chkRandom_Click
' Description:  Whenever the checkbox is changed, show or hide the options as needed.
'

    Dim C As Integer
    Dim RT As String
    Dim I As Integer
    
    RandomGen = (chkRandom.Value = vbChecked)
    If RandomGen Then
        RT = Game.RandomTraits & ","
        For I = 0 To 6
            C = InStr(RT, ",")
            If C > 0 Then txtRandom(I).Text = Mid(RT, 1, C - 1)
            RT = Mid(RT, C + 1)
        Next I
    End If
    Me.Height = IIf(RandomGen, TALL_HEIGHT, SHORT_HEIGHT)
    fraTraits.Visible = RandomGen
    
End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Hide the form.  CharacterName and Race remain empty.
'

    Me.Hide

End Sub

Private Sub cmdChangeling_Click()
'
' Name:         cmdChangeling_Click
' Description:  Set the character as Changeling.
'

    SetCharacter gvRaceChangeling

End Sub

Private Sub cmdDemon_Click()
'
' Name:         cmdDemon_Click
' Description:  Set the character as Demon.
'

    SetCharacter gvRaceDemon

End Sub

Private Sub cmdHunter_Click()
'
' Name:         cmdHunter_Click
' Description:  Set the character as Hunter.
'

    SetCharacter gvRaceHunter

End Sub

Private Sub cmdKueiJin_Click()
'
' Name:         cmdKueiJin_Click
' Description:  Set the character as Kuei-Jin.
'

    SetCharacter gvRaceKueiJin

End Sub

Private Sub cmdWerewolf_Click()
'
' Name:         cmdWerewolf_Click
' Description:  Set the character as Werewolf.
'

    SetCharacter gvRaceWerewolf

End Sub

Private Sub cmdVampire_Click()
'
' Name:         cmdVampire_Click
' Description:  Set the character as Vampire.
'

    SetCharacter gvRaceVampire

End Sub

Private Sub cmdMortal_Click()
'
' Name:         cmdMortal_Click
' Description:  Set the character as Mortal.
'

    SetCharacter gvRaceMortal

End Sub

Private Sub cmdMage_Click()
'
' Name:         cmdMage_Click
' Description:  Set the character as Mage.
'

    SetCharacter gvracemage

End Sub

Private Sub cmdFera_Click()
'
' Name:         cmdFera_Click
' Description:  Set the character as Fera.
'

    SetCharacter gvRaceFera

End Sub

Private Sub cmdMummy_Click()
'
' Name:         cmdMummy_Click
' Description:  Set the character as Mummy.
'

    SetCharacter gvRaceMummy

End Sub

Private Sub cmdVarious_Click()
'
' Name:         cmdVarious_Click
' Description:  Set the character as Various.
'

    SetCharacter gvRaceVarious

End Sub

Private Sub cmdWraith_Click()
'
' Name:         cmdWraith_Click
' Description:  Set the character as Wraith.
'

    SetCharacter gvRaceWraith

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Shift the focus to txtName whenever this form becomes active.
'

    txtName.SetFocus

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the setting for random trait generation.
'

    RandomGen = CBool(GetSetting(App.Title, "Settings", "Random Traits", False))

End Sub

Private Sub txtName_GotFocus()
'
' Name:         txtName_GotFocus
' Description:  Select all the text when the name gets focus.
'

    SelectText txtName

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
'
' Name:         txtName_KeyPress
' Description:  If return is pressed, jump to the first race button (Vampire).
'
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdVampire.SetFocus
    End If
    
End Sub

Private Sub txtRandom_GotFocus(Index As Integer)
'
' Name:         txtRandom_GotFocus
' Description:  Select the Text.
'
    SelectText txtRandom(Index)

End Sub
