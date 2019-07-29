VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHistoryEntry 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   210
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   690
      Width           =   2055
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtReason 
      Height          =   975
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ComboBox cboChange 
      Height          =   315
      ItemData        =   "frmHistoryEntry.frx":0000
      Left            =   960
      List            =   "frmHistoryEntry.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtPoints 
      Height          =   315
      Left            =   2445
      TabIndex        =   4
      Text            =   "1"
      Top             =   720
      Width           =   570
   End
   Begin MSComCtl2.UpDown updPoints 
      Height          =   315
      Left            =   3060
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      OrigLeft        =   6840
      OrigTop         =   960
      OrigRight       =   7275
      OrigBottom      =   1245
      Max             =   999
      Min             =   -999
      Orientation     =   1
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      Caption         =   "Warning: the suggested change and reason are only rough guesses.  Make sure they are correct before committing this entry!"
      Height          =   975
      Left            =   3840
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "&Reason"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1245
      Width           =   615
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "&Change"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   615
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "&Date"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "frmHistoryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean                      'whether or not to make the entry
Private NewDate As Date                         'new entry date
Private NewReason As String                     'new entry reason
Private NewChange As Single                     'new entry change
Private NewChangeType As ExperienceChangeType   'new entry changetype

Public Sub MakeEntry(Experience As ExperienceClass, NewEntry As Boolean, Caption As String, _
                     Optional Reason As String = "", Optional CType As ExperienceChangeType = ecSpent, _
                     Optional Cost As Single = 0, Optional Guess As Boolean = False)
'
' Name:         GetEntry
' Parameters:   Experience          the ExperienceClass to add to
'               Caption             the title of the dialog
' Description:  Get a history entry from the user.
'

    Dim I As Integer

    Load Me

    Me.Caption = Caption

    cboChange.AddItem "Spend"
    cboChange.ItemData(cboChange.NewIndex) = ecSpent
    cboChange.AddItem "Earn"
    cboChange.ItemData(cboChange.NewIndex) = ecEarned
    cboChange.AddItem "Unspend"
    cboChange.ItemData(cboChange.NewIndex) = ecUnspent
    cboChange.AddItem "Lose"
    cboChange.ItemData(cboChange.NewIndex) = ecDeducted
    cboChange.AddItem "Set Earned to"
    cboChange.ItemData(cboChange.NewIndex) = ecSetEarned
    cboChange.AddItem "Set Unspent to"
    cboChange.ItemData(cboChange.NewIndex) = ecSetUnspent
    cboChange.AddItem "Comment"
    cboChange.ItemData(cboChange.NewIndex) = ecComment
    
    With Game.Calendar
        cboDate.Clear
        .First
        Do Until .Off
            cboDate.AddItem Format(.GetGameDate, "mmmm d, yyyy")
            .MoveNext
        Loop
    End With

    If NewEntry Or Experience.Off Then
    
        cboDate.Text = Format(Now, "mmmm d, yyyy")
        cboChange.ListIndex = 0
        If Reason <> "" Then
            I = 0
            Do Until cboChange.ItemData(I) = CType
                I = (I + 1) Mod cboChange.ListCount
                If I = 0 Then Exit Do
            Loop
            cboChange.ListIndex = I
            txtReason.Text = Reason
            txtPoints.Text = CStr(Cost)
            lblGuess.Visible = Guess
        End If
        
    Else
            
        cboDate.Text = Format(Experience.EntryDate, "mmmm d, yyyy")
        I = 0
        Do Until cboChange.ItemData(I) = Experience.EntryChangeType
            I = (I + 1) Mod cboChange.ListCount
            If I = 0 Then Exit Do
        Loop
        cboChange.ListIndex = I
        txtPoints.Text = CStr(Experience.EntryChange)
        txtReason.Text = Experience.EntryReason
        
    End If

    Canceled = False
    
    Me.Show vbModal
    
    If Not Canceled Then
                
        If Not (NewEntry Or Experience.Off) Then Experience.Remove
        
        Experience.Insert NewChange, NewChangeType, NewDate, NewReason
        Game.DataChanged = True
        
    End If
    
End Sub

Private Sub cboChange_Click()
'
' Name:         cboChange_Click
' Description:  Change the visibility of the points controls if "Add Comment" is selected.
'

    txtPoints.Visible = Not (cboChange.ItemData(cboChange.ListIndex) = ecComment)
    updPoints.Visible = txtPoints.Visible

End Sub

Private Sub cboDate_KeyPress(KeyAscii As Integer)
'
' Name:         cboDate_KeyPress
' Parameters:   KeyAscii        code of key pressed
' Description:  Move to the next field.
'

    If KeyAscii = 13 Then
        KeyAscii = 0
        cboChange.SetFocus
    End If

End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Unload this window and allow a new entry.
'

    If IsDate(cboDate.Text) Then
        NewDate = CDate(cboDate.Text)
    Else
        MsgBox "Please enter a valid date.", vbOKOnly + vbExclamation, "Invalid Date"
        Exit Sub
    End If
    NewReason = TrimWhiteSpace(Replace(txtReason.Text, vbCrLf, " "))
    NewChange = Val(txtPoints.Text)
    NewChangeType = cboChange.ItemData(cboChange.ListIndex)
    Canceled = False
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Unload this window and forbid a new entry.
'

    Canceled = True
    Unload Me
    
End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Set the focus when the form appears.
'

    Call cboChange_Click
    cboChange.SetFocus

End Sub

Private Sub txtPoints_GotFocus()
'
' Name:         txtPoints_GotFocus
' Description:  Select the text upon gaining focus.
'

    SelectText txtPoints
    
End Sub

Private Sub txtPoints_KeyPress(KeyAscii As Integer)
'
' Name:         txtPoints_KeyPress
' Parameters:   KeyAscii        code of key pressed
' Description:  Move to the next field.
'

    If KeyAscii = 13 Then
        KeyAscii = 0
        txtReason.SetFocus
    End If

End Sub

Private Sub txtReason_GotFocus()
'
' Name:         txtReason_GotFocus
' Description:  Select the text upon gaining focus.
'

    SelectText txtReason
    
End Sub

Private Sub updPoints_DownClick()
'
' Name:         updPoints_DownClick
' Description:  Decrement the point adjustment.
'

    txtPoints = CStr(Val(txtPoints) - 1)

End Sub

Private Sub updPoints_UpClick()
'
' Name:         updPoints_UpClick
' Description:  Increment the point adjustment.
'

    txtPoints = CStr(Val(txtPoints) + 1)

End Sub


