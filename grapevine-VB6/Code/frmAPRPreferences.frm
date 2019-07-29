VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAPRPreferences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Action and Rumor Settings"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmAPRPreferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdEscClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   6855
      Begin VB.CommandButton cmdDeleteBackground 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddBackground 
         Caption         =   "&Add"
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   2400
         Width           =   1695
      End
      Begin MSComCtl2.UpDown updValue 
         Height          =   285
         Left            =   5535
         TabIndex        =   9
         Top             =   1200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtValue"
         BuddyDispid     =   196613
         OrigLeft        =   5341
         OrigTop         =   1200
         OrigRight       =   5536
         OrigBottom      =   1485
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updPersonal 
         Height          =   285
         Left            =   3015
         TabIndex        =   4
         Top             =   480
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPersonal"
         BuddyDispid     =   196614
         OrigLeft        =   3240
         OrigTop         =   480
         OrigRight       =   3495
         OrigBottom      =   675
         Max             =   9999
         Min             =   -9999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtPersonal 
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.ListBox lstBackgrounds 
         Height          =   1425
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2400
         Width           =   2415
      End
      Begin MSComctlLib.ListView lvwLevels 
         Height          =   1455
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Level"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   1640
         EndProperty
      End
      Begin VB.Label lblLabel 
         Caption         =   "Action &Value"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Backgrounds &with Action Values"
         Height          =   495
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   2445
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Action Values per &Level of Influence or Background"
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   885
         Width           =   2175
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total &Personal Action Value"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   525
         Width           =   2175
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CheckBox chkRumors 
         Caption         =   "Copy all rumor text from the previous session"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   23
         Top             =   2640
         Width           =   3735
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "All additional rumor types from the previous session"
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   22
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Influence rumors"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   21
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Public Knowledge rumors"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   16
         Top             =   480
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Racial rumors (Vampire, Werewolf, etc.)"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   18
         Top             =   960
         Width           =   3855
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Subgroup rumors (sect, auspice, etc.)"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   20
         Top             =   1440
         Width           =   4095
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Group rumors (clan, tribe, etc.)"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   19
         Top             =   1200
         Width           =   4215
      End
      Begin VB.CheckBox chkRumors 
         Caption         =   "Personal rumors"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   17
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Standard Rumors include the following:"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "A&ctions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rumors"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAPRPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Populating As Boolean

Private Const ciPublic = 0
Private Const ciPersonal = 1
Private Const ciRace = 2
Private Const ciGroup = 3
Private Const ciSubgroup = 4
Private Const ciInfluence = 5
Private Const ciPrevious = 6
Private Const ciCopy = 7

Private Sub chkRumors_Click(Index As Integer)
'
' Name:         chkRumors_Click
' Description:  Record the new standard rumor settings.
'

    If Not Populating Then
    
        With Game.APREngine
        
            .PublicRumors = (chkRumors(ciPublic).Value = vbChecked)
            .PersonalRumors = (chkRumors(ciPersonal).Value = vbChecked)
            .RaceRumors = (chkRumors(ciRace).Value = vbChecked)
            .GroupRumors = (chkRumors(ciGroup).Value = vbChecked)
            .SubgroupRumors = (chkRumors(ciSubgroup).Value = vbChecked)
            .InfluenceRumors = (chkRumors(ciInfluence).Value = vbChecked)
            .PreviousRumors = (chkRumors(ciPrevious).Value = vbChecked)
            .CopyPrevious = (chkRumors(ciCopy).Value = vbChecked)
        
        End With
    
        Game.DataChanged = True
    
    End If

End Sub

Private Sub cmdAddBackground_Click()
'
' Name:         cmdAddBackground
' Description:  Add a new background to the list of actionable backgrounds.
'

    Dim NewBackground As String

    NewBackground = InputBox("Enter a background that characters can use for actions:", _
                    "Actionable Background")
    
    NewBackground = Trim(NewBackground)
    
    If NewBackground <> "" Then
    
        lstBackgrounds.AddItem NewBackground
        Game.APREngine.BackgroundActions.Insert NewBackground
        Game.DataChanged = True
        
    End If
    
End Sub

Private Sub cmdDeleteBackground_Click()
'
' Name:         cmdDeleteBackground
' Description:  Delete a background from the list of actionable backgrounds.
'

    If lstBackgrounds.ListIndex > -1 Then
    
        Game.APREngine.BackgroundActions.MoveTo lstBackgrounds.Text
        Game.APREngine.BackgroundActions.RemoveTrait
        lstBackgrounds.RemoveItem lstBackgrounds.ListIndex
        If lstBackgrounds.ListCount > 0 Then lstBackgrounds.ListIndex = 0
        Game.DataChanged = True
        
    End If

End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'

    Unload Me

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Load the APR settings into this form.
'

    Dim I As Integer
    Dim NewItem As ListItem
    
    With Game.APREngine
    
        For I = 1 To 10
            .ActionsPerLevel.MoveTo CStr(I)
            If Not .ActionsPerLevel.Off Then
                Set NewItem = lvwLevels.ListItems.Add(I, , "Influence x" & CStr(I))
                NewItem.ListSubItems.Add , , .ActionsPerLevel.Trait.Total & " Actions"
            End If
        Next I
    
        If lvwLevels.ListItems.Count > 0 Then
            Set lvwLevels.SelectedItem = lvwLevels.ListItems(1)
            Call lvwLevels_ItemClick(lvwLevels.SelectedItem)
        End If
        
        .BackgroundActions.First
        Do Until .BackgroundActions.Off
            lstBackgrounds.AddItem .BackgroundActions.Trait.Name
            .BackgroundActions.MoveNext
        Loop
    
        Populating = True
    
        txtPersonal.Text = .PersonalActions
    
        chkRumors(ciPublic).Value = IIf(.PublicRumors, vbChecked, vbUnchecked)
        chkRumors(ciPersonal).Value = IIf(.PersonalRumors, vbChecked, vbUnchecked)
        chkRumors(ciRace).Value = IIf(.RaceRumors, vbChecked, vbUnchecked)
        chkRumors(ciGroup).Value = IIf(.GroupRumors, vbChecked, vbUnchecked)
        chkRumors(ciSubgroup).Value = IIf(.SubgroupRumors, vbChecked, vbUnchecked)
        chkRumors(ciInfluence).Value = IIf(.InfluenceRumors, vbChecked, vbUnchecked)
        chkRumors(ciPrevious).Value = IIf(.PreviousRumors, vbChecked, vbUnchecked)
        chkRumors(ciCopy).Value = IIf(.CopyPrevious, vbChecked, vbUnchecked)
    
        Populating = False
    
    End With

End Sub

Private Sub lvwLevels_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwLevels_ItemClick
' Description:  Move to a new action level, making it available to edit.
'
    
    With Game.APREngine.ActionsPerLevel
        .MoveToPlace (Item.Index - 1)
        If Not .Off Then
            Populating = True
            txtValue.Text = .Trait.Total
            Populating = False
        End If
    End With

End Sub

Private Sub tabTabs_Click()
'
' Name:         tabTabs_Click()
' Description:  Show the needed frame.
'

    Dim F As Frame
    
    For Each F In fraFrame
        F.Visible = (F.Index = (tabTabs.SelectedItem.Index - 1))
    Next F

End Sub

Private Sub txtPersonal_Change()
'
' Name:         txtPersonal_Change
' Description:  Record the new number of personal actions.
'

    If Not Populating Then
        Game.APREngine.PersonalActions = Val(txtPersonal.Text)
        Game.DataChanged = True
    End If
    
End Sub

Private Sub txtPersonal_GotFocus()
'
' Name:         txtPersonal_GotFocus
' Description:  Select the text upon receiving focus.
'
    SelectText txtPersonal

End Sub

Private Sub txtValue_Change()
'
' Name:         txtValue_Change
' Description:  Record the new action value for this level of influence.
'
    
    If Not (Populating Or lvwLevels.SelectedItem Is Nothing) Then
    
        With Game.APREngine.ActionsPerLevel
            .MoveToPlace (lvwLevels.SelectedItem.Index - 1)
            If Not .Off Then
                .Trait.Total = Val(txtValue.Text)
                lvwLevels.SelectedItem.ListSubItems(1).Text = .Trait.Total & " Actions"
            End If
        End With

    End If

End Sub

Private Sub txtValue_GotFocus()
'
' Name:         txtValue_GotFocus
' Description:  Select the text upon receiving focus.
'
    SelectText txtValue
    
End Sub
