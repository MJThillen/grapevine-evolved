VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Entry"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAddToMenu 
      Caption         =   "&Add this to the Menu"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   2895
   End
   Begin MSComCtl2.UpDown updNumber 
      Height          =   285
      Left            =   855
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtNumber"
      BuddyDispid     =   196611
      OrigLeft        =   735
      OrigTop         =   1200
      OrigRight       =   1230
      OrigBottom      =   1485
      Max             =   9999
      Min             =   -9999
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtNote 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblEntry 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "N&ote (optional):"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblNumber 
      Caption         =   "N&umber:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Name:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   4
      Left            =   3240
      TabIndex        =   11
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "The Custom Entry to Add:"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DescList As LinkedTraitList

Public Sub GetCustom(IsTrait As Boolean, Display As ListDisplayType, AllowAdd As Boolean, _
                     ByRef EntryName As String, ByRef Number As String, ByRef Note As String, _
                     ByRef AddToMenu As Boolean)
'
' Name:         GetCustom
' Parameters:   IsTrait     Whether this is a trait or a field
'               Display     how to display the trait
'               AllowAdd    whether to allow the entry to be added to the menu
'               EntryName   the name of the custom entry
'               Number      the number of the custom entry (if it's a trait)
'               Note        the note for the custom entry
'               AddToMenu   whether to add this to the menu for the session
' Description:  Display the custom entry window.  Display the number controls only
'               if the user is providing a custom trait.
' Returns:      EntryName, Number, Note, AddToMenu
'

    lblNumber.Visible = IsTrait
    txtNumber.Visible = IsTrait
    updNumber.Visible = IsTrait
    txtName.Text = ""
    EntryName = ""
    updNumber.Value = 1
    Number = 1
    txtNote.Text = ""
    Note = ""
    lblEntry.Caption = ""
    chkAddToMenu.Value = vbUnchecked
    chkAddToMenu.Visible = AllowAdd
    
    Me.Show 1, mdiMain
    
    EntryName = Trim(txtName.Text)
    Number = Trim(txtNumber.Text)
    Note = Trim(txtNote.Text)
    AddToMenu = (chkAddToMenu.Value = vbChecked)

End Sub

Private Sub Describe()
'
' Name:         Describe
' Description:  Assemble the custom information and display it in a label.
'

    DescList.Trait.Name = Trim(txtName.Text)
    DescList.Trait.Total = Trim(txtNumber.Text)
    DescList.Trait.Note = Trim(txtNote.Text)
    
    lblEntry.Caption = DescList.DisplayTrait
    
End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Dismiss the window.
'

    txtName.Text = ""
    Me.Hide

End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Verify that the custom data has the correct syntax.  If so,
'               dismiss the window.
'

    If txtName.Text Like "* (*)" Or txtName.Text Like "* x*#" Then
        MsgBox "Parentheses ""()"" and multipliers "" x"" followed by numbers are " & _
                "reserved patterns.  Please choose a different format for the name of your" & _
                "custom entry.", vbOKOnly, "Pattern conflict"
    Else
        Me.Hide
    End If
    
End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Set focus to the entry text field.
'

    txtName.SetFocus

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the one-trait descriptive list.
'

    Set DescList = New LinkedTraitList
    
    DescList.Append "", "1", ""
    DescList.First

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Clean up afterward.
'

    Set DescList = Nothing

End Sub

Private Sub txtName_Change()
'
' Name:         txtName_Change
' Description:  Call Describe to update the lblEntry.
'

    Call Describe

End Sub

Private Sub txtName_GotFocus()
'
' Name:         txtName_GotFocus
' Description:  Select the text in the control.
'

    SelectText txtName

End Sub

Private Sub txtNote_Change()
'
' Name:         txtNote_Change
' Description:  Call Describe to update the lblEntry.
'

    Call Describe

End Sub

Private Sub txtNote_GotFocus()
'
' Name:         txtNote_GotFocus
' Description:  Select the text in the control.
'

    SelectText txtNote

End Sub

Private Sub txtNumber_Change()
'
' Name:         txtNumber_Change
' Description:  Ensure the custom number doesn't go out of the
'               permissible range.
'

    Call Describe

End Sub

Private Sub txtNumber_GotFocus()
'
' Name:         txtNumber_GotFocus
' Description:  Select the text in the control.
'

    SelectText txtNumber

End Sub
