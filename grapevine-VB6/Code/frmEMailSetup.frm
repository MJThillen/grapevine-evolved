VERSION 5.00
Begin VB.Form frmEMailSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-Mail Setup"
   ClientHeight    =   3585
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6360
   Icon            =   "frmEMailSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test E-Mail Settings"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox chkRemember 
      Caption         =   "&Remember Password"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account E-Mail &Address"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   4
      Top             =   1005
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   $"frmEMailSetup.frx":058A
      Height          =   855
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "SMTP &Port"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   645
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "SMTP &Server"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&User Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1365
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Pass&word"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1725
      Width           =   1695
   End
End
Attribute VB_Name = "frmEMailSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SET_USERNAME = 0
Private Const SET_PASSWORD = 1
Private Const SET_SERVER = 2
Private Const SET_PORT = 3
Private Const SET_ADDRESS = 4

Private Canceled As Boolean

Public Sub ShowSetup()
'
' Name:         ShowSetup
' Parameters:   Mailer      SMTPControl to set up
' Description:  Show the E-Mail Setup dialog, allowing the user to specify her
'               SMTP Server, Username and Password
'

' Include the SMTP.ocx control in the project and add an instance to mdiMain,
' naming it smtpMailer, in order to re-enable the code below. Note that SMTP.ocx
' is detected (inaccurately) as a virus component by many antivirus programs.

'    mdiMain.InitializeSMTP
'
'    With mdiMain.smtpMailer
'        txtSettings(SET_SERVER).Text = .Server
'        txtSettings(SET_PORT).Text = CStr(.Port)
'        txtSettings(SET_ADDRESS).Text = .MailFrom
'        txtSettings(SET_USERNAME).Text = .Username
'        txtSettings(SET_PASSWORD).Text = .Password
'        chkRemember.Value = CInt(GetSetting(App.Title, "EMail", "Remember", vbUnchecked))
'    End With
'
'    Me.Show vbModal, mdiMain
'
'    If Not Canceled Then
'
'        With mdiMain.smtpMailer
'            .Server = txtSettings(SET_SERVER).Text
'            .Port = Val(txtSettings(SET_PORT).Text)
'            .MailFrom = txtSettings(SET_ADDRESS).Text
'            .Username = txtSettings(SET_USERNAME).Text
'            .Password = txtSettings(SET_PASSWORD).Text
'
'            SaveSetting App.Title, "EMail", "Server", .Server
'            SaveSetting App.Title, "EMail", "Port", CStr(.Port)
'            SaveSetting App.Title, "EMail", "Address", .MailFrom
'            SaveSetting App.Title, "EMail", "Username", .Username
'
'            If chkRemember.Value = vbChecked And .Username <> "" Then
'                Dim Password As String
'                Password = XORScramble(.Username, XORScramble(xeKey, .Password))
'                SaveSetting App.Title, "EMail", "Pwd", Password
'            Else
'                SaveSetting App.Title, "EMail", "Pwd", ""
'            End If
'
'            SaveSetting App.Title, "EMail", "Remember", CStr(chkRemember.Value)
'        End With
'
'    End If
'
'    Unload Me
    
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
    Canceled = False
    Me.Hide
    
End Sub

Private Sub cmdTest_Click()
'
' Name:         cmdTest_Click
' Description:  Test the SMTP settings.
'

    Dim FinishMsg As String

    If txtSettings(SET_ADDRESS).Text Like "*?@?*.?*" Then
    
        MsgBox "E-Mailing Functions have been disabled in this version of Grapevine."
    
    ' Include the SMTP.ocx control in the project and add an instance to this form,
    ' naming it smtpTest, in order to re-enable the code below. Note that SMTP.ocx
    ' is detected (inaccurately) as a virus component by many antivirus programs.
    
'        With smtpTest
'            .Server = txtSettings(SET_SERVER).Text
'            .Port = Val(txtSettings(SET_PORT).Text)
'            .Username = txtSettings(SET_USERNAME).Text
'            .Password = txtSettings(SET_PASSWORD).Text
'            .SendTo = txtSettings(SET_ADDRESS).Text
'            .MailFrom = txtSettings(SET_ADDRESS).Text
'            .MessageSubject = "Grapevine Test Message"
'            .MessageText = "This message confirms that your e-mail setup in Grapevine works properly."
'            Screen.MousePointer = vbHourglass
'            .SendEmail
'            Screen.MousePointer = vbDefault
'            If .Tag = "" Then
'                FinishMsg = "Success!" & vbCrLf & "No errors were encountered sending a test message."
'                MsgBox FinishMsg, vbOKOnly, "Test E-Mail Settings"
'            Else
'                FinishMsg = "Error sending test message:" & vbCrLf & vbCrLf & .Tag
'                If InStr(LCase(.Tag), "auth") Then
'                    FinishMsg = FinishMsg & vbCrLf & vbCrLf & _
'                                "Are you sure your username and password are correct?" & vbCrLf & _
'                                "Are you sure you need them at all?"
'                End If
'                MsgBox FinishMsg, vbOKOnly + vbExclamation, "Test E-Mail Settings"
'            End If
'        End With

    Else
        MsgBox "Invalid E-Mail Address.", vbExclamation, "Invalid Address"
    End If

End Sub

' Include the SMTP.ocx control in the project and add an instance to this form,
' naming it smtpTest, in order to re-enable the code below. Note that SMTP.ocx
' is detected (inaccurately) as a virus component by many antivirus programs.

'Private Sub smtpTest_ErrorSMTP(ByVal Number As Integer, Description As String)
''
'' Name:         smtpTest_ErrorSMTP
'' Description:  Set the tag (error field) on a bad send.
''
'    smtpTest.Tag = TrimWhiteSpace(Description)
'
'End Sub
'
'Private Sub smtpTest_SendSMTP()
''
'' Name:         smtpTest_SendSMTP
'' Description:  Clear the tag (error field) on a successful send.
''
'    smtpTest.Tag = ""
'
'End Sub

Private Sub txtSettings_GotFocus(Index As Integer)
'
' Name:         txtSettings_GotFocus
' Description:  Select text.
'
    SelectText txtSettings(Index)

End Sub

Private Sub txtSettings_KeyPress(Index As Integer, KeyAscii As Integer)
'
' Name:         txtSettings_KeyPress
' Description:  Move to next field on vbCr.
'
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub
