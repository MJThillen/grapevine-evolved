VERSION 5.00
Begin VB.Form frmDuplicate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Duplicate Entry"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CheckBox chkAll 
      Caption         =   "Perform this action for &all duplicates"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton cmdKeepOlder 
      Caption         =   "Keep &Older Version"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdKeepNewer 
      Caption         =   "Keep &Newer Version"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdRenameIncoming 
      Caption         =   "Rename &Incoming Entry"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdRenameExisting 
      Caption         =   "Rename &Existing Entry"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtName 
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
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdSkip 
      Cancel          =   -1  'True
      Caption         =   "&Skip"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblIncoming 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      Caption         =   "The incoming version was last modified:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      Caption         =   "The existing version was last modified:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label lblExisting 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      Caption         =   "What do you want to do about it?"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmDuplicate.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Caption         =   "You have tried to load an entry with the above name into the game database, but an entry already exists under this name."
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "frmDuplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum DupAction
    RenameNew
    RenameOld
    ReplaceOld
    Skip
    keepolder
    KeepNewer
End Enum

Public Action As DupAction          'the action selected by the user
Public All As Boolean               'whether to perform that action on all duplicates
Public NewName As String            'new name entered by user
Private OriginalName As String      'original duplicate name
Private SearchList As LinkedList    'the list of existing players or characters

Public Sub FixDuplicate(Name As String, List As LinkedList, ExMod As Date, InMod As Date, Rename As Boolean)
'
' Name:         FixDuplicate
' Parameters:   Name    the duplicate name
'               List    the list of existing players or characters
'               ExMod   Last Modified date of the exisiting entity
'               InMod   Last Modified date of the incoming entity
'               Rename  Whether or not renaming is a viable option
' Description:  Initialize the window display and data, then show the window
'

    Select Case InMod
        Case Is < ExMod
            lblExisting.Caption = CStr(ExMod) & " (Newer)"
            lblIncoming.Caption = CStr(InMod) & " (Older)"
        Case Is = ExMod
            lblExisting.Caption = CStr(ExMod)
            lblIncoming.Caption = "At the same time."
        Case Is > ExMod
            lblExisting.Caption = CStr(ExMod) & " (Older)"
            lblIncoming.Caption = CStr(InMod) & " (Newer)"
    End Select
    
    cmdKeepNewer.Visible = (ExMod <> InMod)
    cmdKeepOlder.Visible = (ExMod <> InMod)

    txtName.Locked = Not Rename
    cmdRenameExisting.Visible = Rename
    cmdRenameIncoming.Visible = Rename

    OriginalName = Name
    Set SearchList = List
    txtName.Text = Name
    txtName.SelStart = 0
    txtName.SelLength = Len(Name)
    chkAll.Value = vbUnchecked
    Me.Show vbModal, mdiMain
        
End Sub

Private Sub chkAll_Click()
'
' Name:         chkAll_Click
' Description:  Select whether the action should be carried out for all
'               duplicates.
'
    All = (chkAll.Value = vbChecked)

End Sub

Private Sub cmdKeepNewer_Click()
'
' Name:         cmdKeepNewer_Click
' Description:  signify that the user wants to keep the newer data.
'
    Action = KeepNewer
    Me.Hide

End Sub

Private Sub cmdKeepOlder_Click()
'
' Name:         cmdKeepOlder_Click
' Description:  signify that the user wants to keep the Older data.
'
    Action = keepolder
    Me.Hide

End Sub

Private Sub cmdRenameExisting_Click()
'
' Name:         cmdRenameExisting_Click
' Description:  If the name entered is not in use, signify that the user
'               wants to change the name that already exists.
'
    
    txtName = Trim(txtName)
    
    If txtName = "" Then Exit Sub

    If txtName = OriginalName Then
        MsgBox "Please type a new name in the text box.", vbOKOnly, "Rename"
        Exit Sub
    End If

    SearchList.MoveTo txtName
    If SearchList.Off Then
        Action = RenameOld
        NewName = txtName
        Me.Hide
    Else
        MsgBox "That name is also in use.", vbExclamation + vbOKOnly, "Ditto!"
    End If

End Sub

Private Sub cmdRenameIncoming_Click()
'
' Name:         cmdRenameIncoming_Click
' Description:  If the name entered is not in use, signify that the user
'               wants to change the name of the incoming data.
'
    
    txtName = Trim(txtName)
    
    If txtName = "" Then Exit Sub

    If txtName = OriginalName Then
        MsgBox "Please type a new name in the text box.", vbOKOnly, "Rename"
        Exit Sub
    End If
    
    SearchList.MoveTo txtName
    If SearchList.Off Then
        Action = RenameNew
        NewName = txtName
        Me.Hide
    Else
        MsgBox "That name is also in use.", vbExclamation + vbOKOnly, "Ditto!"
    End If

End Sub

Private Sub cmdReplace_Click()
'
' Name:         cmdReplace_Click
' Description:  Signify that the user wants to replace the existing data
'               with the incoming data.
'

    Action = ReplaceOld
    Me.Hide

End Sub

Private Sub cmdSkip_Click()
'
' Name:         cmdSkip_Click
' Description:  Signify that the user wants to skip the incoming data.
'

    Action = Skip
    Me.Hide
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the window's data.
'

    All = False

End Sub
