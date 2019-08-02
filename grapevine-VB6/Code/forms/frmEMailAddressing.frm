VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEMailAddressing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-Mail Addressing"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmEMailAddressing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRecipients 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   5655
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ß"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.ListBox lstAddresses 
         Height          =   1815
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   2535
      End
      Begin VB.ListBox lstPlayers 
         Height          =   1815
         Left            =   3120
         TabIndex        =   17
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "à"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   1815
         Width           =   375
      End
      Begin VB.CheckBox chkSelected 
         Caption         =   "Send to &Players of Selected Characters"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "&Potential Recipients"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   16
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblRecipients 
         Alignment       =   2  'Center
         Caption         =   "Additional R&ecipients"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame fraAttachments 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   1320
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdUnattach 
         Caption         =   "Remo&ve"
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "A&dd..."
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.ListBox lstAttachments 
         Height          =   2205
         Left            =   0
         TabIndex        =   21
         Top             =   345
         Width           =   3135
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Attachments"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.TextBox txtHeader 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   25
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   24
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   3135
   End
   Begin MSComctlLib.TabStrip tabRecipients 
      Height          =   3135
      Left            =   1200
      TabIndex        =   9
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5530
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Send &To"
            Key             =   "to"
            Object.Tag             =   "Send"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&CC"
            Key             =   "cc"
            Object.Tag             =   "CC"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&BCC"
            Key             =   "bcc"
            Object.Tag             =   "BCC"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Message"
            Key             =   "message"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Attachments"
            Key             =   "attach"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Sending"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   855
   End
   Begin VB.Label lblSending 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblFrom 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "From"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   2
      Top             =   630
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Recipients/ Content"
      Height          =   420
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1845
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Reply-To"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1005
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Subject"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1365
      Width           =   855
   End
End
Attribute VB_Name = "frmEMailAddressing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Subject = 1
Private Const ReplyTo = 0

Private Const adSENDTO = 1
Private Const adCC = 2
Private Const adBCC = 3
Private Const adMESSAGE = 4
Private Const adATTACH = 5

Private Const OTHER_ADDRESS = "(Specify Address...)"
Private Const HEADER_CAPTION = "Enter a message here that appears at the beginning of the e-mail."

Private Canceled As Boolean
Private CurrentTab As Integer
Private Addresses(1 To 3) As String
Private MultiMessage As Boolean
Private Attachments As Collection
Private Populating As Boolean

' Include the SMTP.ocx control in the project in order to re-enable the code
' below. Note that SMTP.ocx is detected (inaccurately) as a virus component
' by many antivirus programs.

'Public Sub ShowAddressing(Mailer As SMTP, Content As Long, Description As String)
''
'' Name:         ShowAddressing
'' Parameters:   Mailer      SMTP control being edited
''               Content     Content of the item(s) to be mailed
''               Description What is being sent
'' Description:  Initialize the form, show it, and process its content upon close.
''
'
'    Dim I As Integer
'
'    Set Attachments = New Collection
'
'    lblSending = Description
'    lblFrom.Caption = Mailer.MailFrom
'    txtInput(Subject).Text = OutputEngine.MessageSubject
'    txtInput(ReplyTo).Text = OutputEngine.ReplyTo
'
'    MultiMessage = ooReport And Not Content
'    If MultiMessage Then MultiMessage = Content And (ooCharacters Or ooPlayers)
'    chkSelected.Visible = MultiMessage
'    chkSelected.Tag = IIf(Content And ooPlayers, " to &Players Selected for Output", _
'                              " to &Players of Characters Selected for Output")
'    chkSelected.Caption = "Send" & chkSelected.Tag
'
'    txtHeader.Text = OutputEngine.MessageHeader
'    If txtHeader.Text = "" Then txtHeader.Text = HEADER_CAPTION
'
'    Addresses(adSENDTO) = OutputEngine.SendTo
'    Addresses(adCC) = OutputEngine.CC
'    Addresses(adBCC) = OutputEngine.BCC
'
'    lstPlayers.Clear
'    lstPlayers.AddItem OTHER_ADDRESS
'    With Game.QueryEngine.QueryList
'        .First
'        Do Until .Off
'            If .Item.Inventory = qiPlayers Then
'                lstPlayers.AddItem "(" & .Item.Name & ")"
'            End If
'            .MoveNext
'        Loop
'    End With
'    PlayerList.First
'    Do Until PlayerList.Off
'        If PlayerList.Item.EMail <> "" Then lstPlayers.AddItem PlayerList.Item.Name
'        PlayerList.MoveNext
'    Loop
'    With Game.QueryEngine.QueryList
'        .First
'        Do Until .Off
'            If .Item.Inventory = qiCharacters Then
'                lstPlayers.AddItem "(" & .Item.Name & ")"
'            End If
'            .MoveNext
'        Loop
'    End With
'
'    RefreshAddresses adSENDTO
'
'    CurrentTab = adSENDTO
'    Set tabRecipients.SelectedItem = tabRecipients.Tabs(adSENDTO)
'
'    Me.Show vbModal, mdiMain
'
'    If Not Canceled Then
'
'        OutputEngine.SendTo = Addresses(adSENDTO)
'        OutputEngine.CC = Addresses(adCC)
'        OutputEngine.BCC = Addresses(adBCC)
'        OutputEngine.ReplyTo = txtInput(ReplyTo).Text
'        OutputEngine.MessageSubject = txtInput(Subject).Text
'        OutputEngine.MessageHeader = txtHeader.Text
'        If txtHeader.Text = HEADER_CAPTION Then OutputEngine.MessageHeader = ""
'        Mailer.Tag = ""
'        SaveSetting App.Title, "EMail", "Reply-To", txtInput(ReplyTo).Text
'        For I = 1 To Attachments.Count
'            Mailer.Attachments.Add Attachments.Item(I)
'        Next I
'
'    Else
'
'        Mailer.Tag = "Canceled"
'
'    End If
'
'    Set Attachments = Nothing
'
'End Sub
'
Private Sub RefreshAddresses(Addex As Integer)
'
' Name:         RefreshAddresses
' Parameters:   Addex       index to comma-delimited list of addresses
' Description:  Fill lstAddresses with the addresses from the given string
'

    Dim Adlist As String
    Dim Address As String
    Dim LastComma As Long
    Dim I As Long
    
    Populating = True
    
    lstAddresses.Clear
    chkSelected.Value = vbUnchecked
    
    If Addresses(Addex) <> "" Then
    
        Adlist = Addresses(Addex) & ","
        LastComma = 1
        I = InStr(Adlist, ",")
        
        Do Until I = 0
            If OutsideQuotes(Adlist, I) Then
                Address = Mid(Adlist, LastComma, I - LastComma)
                If Address = SendToSelect Then
                    chkSelected.Value = vbChecked
                Else
                    lstAddresses.AddItem Address
                End If
                LastComma = I + 1
            End If
            I = InStr(I + 1, Adlist, ",")
        Loop

    End If

    Populating = False

    RefreshRecipientLabel

End Sub

Private Sub RefreshRecipientLabel()
'
' Name:         RefreshRecipientLabel
' Description:  Update the lblRecipients and chkSelected controls according to current selections.
'

    Dim Cap As String
    
    Cap = "R&ecipients"
    Select Case CurrentTab
        Case adCC:  Cap = "CC " & Cap
        Case adBCC:  Cap = "BCC " & Cap
    End Select
    If chkSelected.Value = vbChecked And MultiMessage Then Cap = "Additional " & Cap
    lblRecipients.Caption = Cap
    
End Sub

Private Sub chkSelected_Click()
'
' Name:         chkSelected_Click
' Description:  Update the lblRecipient control, adjust the SendToSelected component of the address.
'

    If Not Populating Then

        If chkSelected.Value = vbChecked Then
            Addresses(CurrentTab) = SendToSelect & _
                                    IIf(Addresses(CurrentTab) = "", "", "," & Addresses(CurrentTab))
        Else
            Addresses(CurrentTab) = Replace(Addresses(CurrentTab), SendToSelect & ",", "")
            Addresses(CurrentTab) = Replace(Addresses(CurrentTab), SendToSelect, "")
        End If
    
        RefreshRecipientLabel
    
    End If
    
End Sub

Private Sub cmdAdd_Click()
'
' Name:         cmdAdd_Click
' Description:  Add an address to the list.
'

    Dim NewEMail As String
    
    If lstPlayers.ListIndex > -1 Then
    
        If lstPlayers.Text = OTHER_ADDRESS Then
            NewEMail = InputBox("Enter an e-mail address:", "Add E-Mail Address")
            If Not NewEMail = "" Then
                If Not (NewEMail Like "*?@?*.?*" And Not NewEMail Like "*,*") Then
                    MsgBox NewEMail & " is not a valid e-mail address.", vbOKOnly, "Invalid Address"
                    Exit Sub
                End If
            End If
        ElseIf Left(lstPlayers.Text, 1) = "(" And Right(lstPlayers.Text, 1) = ")" Then
        
            NewEMail = Mid(lstPlayers.Text, 2, Len(lstPlayers.Text) - 2)
            With Game.QueryEngine
                .QueryList.MoveTo NewEMail
                If Not .QueryList.Off Then
                
                    .MakeQuery .QueryList.Item
                    .Results.First
                    Do Until .Results.Off
                        NewEMail = ""
                        If .QueryList.Item.Inventory = qiPlayers Then
                            If .Results.Item.EMail <> "" Then NewEMail = .Results.Item.ExpandedEMail
                        Else
                            PlayerList.MoveTo .Results.Item.Player
                            If Not PlayerList.Off Then
                                If PlayerList.Item.EMail <> "" Then NewEMail = PlayerList.Item.ExpandedEMail
                            End If
                        End If
                        If Not (NewEMail = "" Or InStr(Addresses(CurrentTab), NewEMail) > 0) Then
                            Addresses(CurrentTab) = Addresses(CurrentTab) & "," & NewEMail
                            lstAddresses.AddItem NewEMail
                        End If
                        .Results.MoveNext
                    Loop
                    
                    lstAddresses.ListIndex = lstAddresses.NewIndex
                    
                    If Left(Addresses(CurrentTab), 1) = "," Then
                        Addresses(CurrentTab) = Mid(Addresses(CurrentTab), 2)
                    End If
                    
                End If
            End With
            NewEMail = ""
        
        Else
            PlayerList.MoveTo lstPlayers.Text
            If Not PlayerList.Off Then
                NewEMail = PlayerList.Item.ExpandedEMail
            End If
        End If
        
        If Not (NewEMail = "" Or InStr(Addresses(CurrentTab), NewEMail) > 0) Then
            lstAddresses.AddItem NewEMail
            lstAddresses.ListIndex = lstAddresses.NewIndex
            If Addresses(CurrentTab) <> "" Then Addresses(CurrentTab) = Addresses(CurrentTab) & ","
            Addresses(CurrentTab) = Addresses(CurrentTab) & NewEMail
        End If
        
    End If

End Sub

Private Sub cmdAttach_Click()
'
' Name:         cmdAttach_Click
' Description:  Browse for and attach a file to the message.
'

    With mdiMain.cmnDialog
        .DialogTitle = "Attach"
        .InitDir = GetSetting(App.Title, "Files", "GameDir", CurDir)
        .FileName = ""
        .Flags = cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNHideReadOnly
        .Filter = "Grapevine Game Files|*.gv3;*.gv2|Grapevine Exchange Files|*.gex|" & _
                  "Grapevine Menu Files|*.gvm|All Files|*.*"
        .FilterIndex = 1
        
        On Error GoTo cmdAttach_Finish
        .ShowOpen
        On Error GoTo 0
        
        If Attachments.Count = 0 Then
            Attachments.Add .FileName
            lstAttachments.AddItem .FileTitle
        Else
            Attachments.Add .FileName, after:=Attachments.Count
            lstAttachments.AddItem .FileTitle, lstAttachments.ListCount
        End If
    
    End With
    
cmdAttach_Finish:

End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Cancel out of the setup screen.
'
    Canceled = True
    Me.Hide

End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Commit the values from the setup screen (by returning to ShowSetup).
'
    
    If Addresses(adSENDTO) = "" Then
        MsgBox "You must specify at least one recipient.", vbOKOnly + vbExclamation, "Send E-Mail"
    Else
        Canceled = False
        Me.Hide
    End If
    
End Sub

Private Sub cmdRemove_Click()
'
' Name:         cmdRemove_Click
' Description:  Remove an address from the list.
'

    Dim StorePos As Long
    Dim Index As Long
    
    StorePos = lstAddresses.ListIndex
    
    If Not StorePos = -1 Then
        Addresses(CurrentTab) = Replace(Addresses(CurrentTab), lstAddresses.Text & ",", "")
        Addresses(CurrentTab) = Replace(Addresses(CurrentTab), lstAddresses.Text, "")
        lstAddresses.RemoveItem lstAddresses.ListIndex
        If StorePos >= lstAddresses.ListCount Then StorePos = StorePos - 1
        lstAddresses.ListIndex = StorePos
    End If

End Sub

Private Sub cmdUnattach_Click()
'
' Name:         cmdUnattach_Click
' Description:  Remove an attachment from the message.
'
    
    Dim I As Integer
    
    I = lstAttachments.ListIndex
    If I > -1 Then
        Attachments.Remove I + 1
        lstAttachments.RemoveItem I
        If I >= lstAttachments.ListCount Then I = lstAttachments.ListCount - 1
        lstAttachments.ListIndex = I
    End If

End Sub

Private Sub lstAddresses_DblClick()
'
' Name:         lstAddresses_DblClick
' Description:  Remove an address from the address list.
'
    Call cmdRemove_Click

End Sub

Private Sub lstAddresses_KeyPress(KeyAscii As Integer)
'
' Name:         lstAddresses_KeyPress
' Description:  Remove an address from the address list.
'
    If KeyAscii = vbKeySpace Or KeyAscii = vbKeyDelete Then Call cmdRemove_Click

End Sub

Private Sub lstAttachments_DblClick()
'
' Name:         lstAttachments_DblClick
' Description:  Remove an attachment.
'
    Call cmdUnattach_Click
    
End Sub

Private Sub lstPlayers_DblClick()
'
' Name:         lstPlayers_DblClick
' Description:  Add a player to the address list.
'
    Call cmdAdd_Click
    
End Sub

Private Sub lstPlayers_KeyPress(KeyAscii As Integer)
'
' Name:         lstPlayers_KeyPress
' Description:  Add a player to the address list.
'
    If KeyAscii = vbKeySpace Then Call cmdAdd_Click

End Sub

Private Sub tabRecipients_Click()
'
' Name:         tabRecipients
' Description:  Refresh the address list.
'

    CurrentTab = tabRecipients.SelectedItem.Index
    
    txtHeader.Visible = (CurrentTab = adMESSAGE)
    fraAttachments.Visible = (CurrentTab = adATTACH)
    fraRecipients.Visible = Not (txtHeader.Visible Or fraAttachments.Visible)
    If fraRecipients.Visible Then
        RefreshAddresses CurrentTab
        chkSelected.Caption = tabRecipients.SelectedItem.Tag & chkSelected.Tag
    End If
    
End Sub

Private Sub txtHeader_GotFocus()
'
' Name:         txtHeader_GotFocus
' Description:  Select the text.
'
    SelectText txtHeader
    cmdOK.Default = False
    
End Sub

Private Sub txtHeader_LostFocus()
'
' Name:         txtHeader_GotFocus
' Description:  Make the Send button default again.
'
    cmdOK.Default = True
    
End Sub
