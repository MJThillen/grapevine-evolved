VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Grapevine"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5730
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5730
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtWebPage 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Image imgGrapevineIcon 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   240
         Picture         =   "frmAbout.frx":058A
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0E54
         ForeColor       =   &H00000000&
         Height          =   1170
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Width           =   3885
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   3885
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   960
         TabIndex        =   3
         Top             =   1020
         Width           =   3885
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Image imgPortrait 
         BorderStyle     =   1  'Fixed Single
         Height          =   3450
         Index           =   0
         Left            =   120
         Tag             =   $"frmAbout.frx":0F76
         ToolTipText     =   "Click me!"
         Top             =   240
         Width           =   2070
      End
      Begin VB.Image imgPortrait 
         BorderStyle     =   1  'Fixed Single
         Height          =   3450
         Index           =   1
         Left            =   120
         Tag             =   "(For some reason I felt like I had to include a real picture of me somewhere too!)"
         Top             =   240
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label lblMe 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":1082
         Height          =   1575
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblMe 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":118E
         Height          =   1215
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   1800
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   3480
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtCopyright 
         Height          =   2535
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmAbout.frx":123B
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtCredits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   2775
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmAbout.frx":330B
         Top             =   120
         Width           =   4935
      End
   End
   Begin MSComctlLib.TabStrip tabTabStrip 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grapevine"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Copyright © 2003"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Adam Cerling"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "-- with &Thanks"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Dismiss the window.
'

    Unload Me

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Populate the program version and contact information
'               when the window loads.
'

    lblTitle = App.Title
    lblVersion = "Version " & App.Major & "." & App.Minor & ", Revision " & App.Revision
    If App.Major = 1 Then lblVersion = "Version 1.99 " & App.Comments
    txtWebPage = URLMainPage
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Parameters:   Cancel          set to cancel the unload
' Description:  Dismiss the window.
'

    Unload Me

End Sub

Private Sub imgPortrait_Click(Index As Integer)
'
' Name:         imgPortrait_Click
' Description:  Switch the pictures.
'

    imgPortrait(Index).Visible = False
    
    Index = (Index + 1) Mod imgPortrait.Count
    
    imgPortrait(Index).Visible = True
    
    lblMe(0).Caption = imgPortrait(Index).Tag

End Sub

Private Sub tabTabStrip_Click()
'
' Name:         tabTabStrip_Click
' Description:  Display correct frame.
'

    If Not fraTab(tabTabStrip.SelectedItem.Index - 1).Visible Then
        
        Dim fTab As Frame
        For Each fTab In fraTab
            fTab.Visible = (fTab.Index = tabTabStrip.SelectedItem.Index - 1)
        Next fTab
        
    End If

End Sub

Private Sub txtWebPage_Click()
'
' Name:         txtWebPage_Click
' Description:  Launch the user's browser, loading the Grapevine web page.
'

    mdiMain.LaunchBrowser txtWebPage.Text

End Sub
