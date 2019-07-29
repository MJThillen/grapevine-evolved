VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOutputRoster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print or Export Rosters"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   8910
   Icon            =   "frmOutputRoster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboDate 
      Height          =   315
      ItemData        =   "frmOutputRoster.frx":058A
      Left            =   5160
      List            =   "frmOutputRoster.frx":058C
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   720
      Width           =   3255
   End
   Begin VB.ComboBox cboRoster 
      Height          =   315
      ItemData        =   "frmOutputRoster.frx":058E
      Left            =   480
      List            =   "frmOutputRoster.frx":059B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox rtfEasel 
      Height          =   495
      Left            =   5400
      TabIndex        =   29
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmOutputRoster.frx":05D1
   End
   Begin VB.OptionButton optPlainText 
      Caption         =   "Plain Text"
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optRTF 
      Caption         =   "Rich Text Format (RTF)"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   1560
      Value           =   -1  'True
      Width           =   2055
   End
   Begin MSComCtl2.UpDown updCopies 
      Height          =   285
      Left            =   5700
      TabIndex        =   15
      Top             =   2520
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtCopies"
      BuddyDispid     =   196613
      OrigLeft        =   5880
      OrigTop         =   960
      OrigRight       =   6135
      OrigBottom      =   1155
      Max             =   99
      Min             =   1
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtCopies 
      Height          =   285
      Left            =   5160
      TabIndex        =   14
      Text            =   "1"
      Top             =   2520
      Width           =   540
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "Printer Set&up..."
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectActive 
      Caption         =   """Acti&ve"" Status"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdExportRoster 
      Caption         =   "&Export"
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrintRoster 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&All"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearSelect 
      Caption         =   "&Clear Selections"
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ListBox lstMembers 
      Columns         =   2
      Height          =   2595
      Left            =   480
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   5280
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   4800
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      Filter          =   "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date"
      Height          =   195
      Index           =   6
      Left            =   5040
      TabIndex        =   31
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lblLabels 
      Caption         =   "Enter or Select the Appropriate &Date:"
      Height          =   195
      Index           =   1
      Left            =   5160
      TabIndex        =   8
      Top             =   420
      Width           =   3300
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Select a Roster Type:"
      Height          =   195
      Index           =   10
      Left            =   480
      TabIndex        =   1
      Top             =   420
      Width           =   4020
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Roster Type"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   870
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Format"
      Height          =   195
      Index           =   11
      Left            =   5040
      TabIndex        =   10
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   12
      Left            =   4920
      TabIndex        =   28
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Copies:"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   13
      Top             =   2280
      Width           =   525
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0%"
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   4800
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Export to File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   5040
      TabIndex        =   25
      Top             =   3600
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   7
      Left            =   4920
      TabIndex        =   24
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label lblListCaption 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   4020
   End
   Begin VB.Label lblLabels 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5040
      TabIndex        =   23
      Top             =   2040
      Width           =   405
   End
   Begin VB.Label lblRosterCaption 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   4455
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   4
      Left            =   4920
      TabIndex        =   22
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblMeter 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   9
      Left            =   240
      TabIndex        =   30
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   13
      Left            =   4920
      TabIndex        =   32
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmOutputRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Active = 1
Const Other = 0

Private MemberList As LinkedList    'Linked list of characters or players
Private InitRoster As RosterType    'Type of roster to generate

Public Sub ShowRoster(Which As RosterType)
'
' Name:         ShowRoster
' Parameters:   Which       the kind of roster to show
' Description:  Load the window, displaying the appropriate roster.
'

    InitRoster = Which
    Me.Show vbModal

End Sub

Private Function RefreshControls(WhichRoster As RosterType)
'
' Name:         RefreshControls
' Parameters:   Which   the kind of roster for which to refresh
' Description:  Display the controls and lists needed to generate the given roster.
'

    lstMembers.Clear

    If WhichRoster = roPlayer Then
        
        Set MemberList = PlayerList
        PlayerList.First
        Do Until PlayerList.Off
            lstMembers.AddItem PlayerList.Item.Name
            PlayerList.MoveNext
        Loop
        cmdSelectAll.Caption = "&All Players"
        cmdSelectActive.Enabled = False
        
    Else
    
        Set MemberList = CharacterList
        CharacterList.First
        Do Until CharacterList.Off
            lstMembers.AddItem CharacterList.Item.Name
            Select Case CharacterList.Item.Status
                Case ActiveStatus
                    lstMembers.ItemData(lstMembers.NewIndex) = Active
                Case Else
                    lstMembers.ItemData(lstMembers.NewIndex) = Other
            End Select
            CharacterList.MoveNext
        Loop
        cmdSelectAll.Caption = "&All Characters"
        cmdSelectActive.Enabled = True
        
    End If

    SelectMembers True

    lblRosterCaption = cboRoster.List(WhichRoster)

    If WhichRoster = roAttendance Then
    
        lblListCaption = "&Select the Characters for whom to Generate a List:"
        cmdPrintRoster.Caption = "&Print List"
        cmdExportRoster.Caption = "&Export List"
    
    Else
    
        cmdPrintRoster.Caption = "&Print Roster"
        cmdExportRoster.Caption = "&Export Roster"
        If WhichRoster = roCharacter Then
            lblListCaption = "&Select the Characters for whom to Generate a Roster:"
        Else
            lblListCaption = "&Select the Players for whom to Generate a Roster:"
        End If
    
    End If

End Function

Private Function CountSelected()
'
' Name:         CountSelected
' Description:  Count the number of selected items in the list.
'

    Dim Entry As Integer
    Dim Total As Integer

    For Entry = 0 To lstMembers.ListCount - 1
        If lstMembers.Selected(Entry) Then Total = Total + 1
    Next Entry
    
    CountSelected = Total

End Function

Private Sub SelectMembers(Maybe As Boolean)
'
' Name:         SelectMembers
' Parameters:   Maybe       whether to select or deselect
' Description:  Select or deselect every item in the list.
'

    Dim Entry As Integer
    
    For Entry = 0 To lstMembers.ListCount - 1
        lstMembers.Selected(Entry) = Maybe
    Next Entry

End Sub

Private Sub cboRoster_Click()
'
' Name:         cboRoster_Click
' Description:  Select a new type of roster.
'

    RefreshControls cboRoster.ListIndex

End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Dismiss this window.
'
    
    Unload Me

End Sub

Private Sub cmdExportRoster_Click()
'
' Name:         cmdExportRoster_Click
' Description:  Prompt the user for a filename, then export the selected roster
'               for in the selected format for the selected players/characters.
'

    Dim DelimNames As String
    
    Dim FileNum As Integer
    Dim TemplateNum As Integer
    Dim TemplateName As String
    
    Dim Entry As Integer
    Dim TotalSelected As Integer
    
    TotalSelected = CountSelected
    If TotalSelected = 0 Then Exit Sub
    
    With cmnDialog
        
        If optPlainText Then
            .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .DefaultExt = "txt"
        Else
            .Filter = "Rich Text Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
            .DefaultExt = "rtf"
        End If
        
        .FileName = cboRoster.Text & "." & .DefaultExt
        .DialogTitle = "Save " & cboRoster.Text & " As..."
        .InitDir = GetSetting(App.Title, "Files", "DefaultDir", CurDir)
        .Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        On Error Resume Next
        .ShowSave
        If Err <> 0 Then Exit Sub 'Canceled
    
    End With
    
    SaveSetting App.Title, "Files", "DefaultDir", CurDir
    
    On Error GoTo ExportError
    
    Screen.MousePointer = vbHourglass
    
    DelimNames = vbCrLf
    
    For Entry = 0 To lstMembers.ListCount - 1
        
        If lstMembers.Selected(Entry) Then DelimNames = DelimNames & lstMembers.List(Entry) & vbCrLf
        MemberList.MoveNext
        
    Next Entry
    
    FileNum = FreeFile
    Open cmnDialog.FileName For Output As #FileNum
    OutputAid.Destination = goFile
    OutputAid.FileLoc = FileNum
    
    If optPlainText Then
        
        OutputAid.SetStandardPageWidth
        OutputAid.OutputTextRoster cboRoster.ListIndex, cboDate.Text, DelimNames, vbCrLf
        
    Else 'optRTF
        
        Select Case cboRoster.ListIndex
            Case roAttendance: TemplateName = AttendanceTemplateName
            Case roCharacter: TemplateName = CharacterRosterTemplateName
            Case roPlayer: TemplateName = PlayerRosterTemplateName
        End Select
        
        TemplateNum = FreeFile
        Open App.Path & "\" & TemplateName For Input As #TemplateNum
        OutputAid.FilterRTFRoster TemplateNum, cboRoster.ListIndex, cboDate.Text, _
                DelimNames, vbCrLf
        Close #TemplateNum
                
    End If
        
    SaveSetting App.Title, "Output", "RTF Format", optRTF
    
    GoTo ExportFinish
    
ExportError:
    MsgBox "Problem saving file: " & Err.Description, vbOKOnly, "File Error"
    Resume ExportFinish

ExportFinish:
    
    Close #FileNum
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdPrintRoster_Click()
'
' Name:         cmdPrintRoster_Click
' Description:  Print the selected roster in the selected format
'               for the selected players/characters.
'

    Dim DelimNames As String
    
    Dim TemplateNum As Integer
    Dim TemplateName As String
    Dim TempFileNum As Integer
    Dim RTFPageLoc As Integer
    
    Dim Entry As Integer
    Dim Cancel As Boolean
    Dim DocPrinting As Boolean
    Dim Percent As Single
    Dim Copy As Integer

    If CountSelected = 0 Then Exit Sub

    txtCopies = CStr(Val(txtCopies))
    If Val(txtCopies) < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    lblProgress = "0%"
    lblProgress.Visible = True
    lblMeter.Width = 0
    lblMeter.Visible = True

    On Error GoTo PrintError
    
    PrinterAct = paSTART

    If optPlainText Then
        PrinterAct = paFONT
        Printer.Font.Name = "Courier New"
        Printer.Font.Size = 12
        PrinterAct = paWIDTH
        OutputAid.SetStandardPageWidth Cancel
        If Cancel Then Exit Sub
        On Error GoTo 0
    End If

    OutputAid.Destination = goPrinter

    DelimNames = vbCrLf
    For Entry = 0 To lstMembers.ListCount - 1
        If lstMembers.Selected(Entry) Then DelimNames = DelimNames & lstMembers.List(Entry) & vbCrLf
        MemberList.MoveNext
    Next Entry
    
    For Copy = 1 To Val(txtCopies)

        DocPrinting = False
        
        If optPlainText Then

            If DocPrinting Then
                PrinterAct = paNEWPAGE
                Printer.NewPage
            End If
            DocPrinting = True
            PrinterAct = "print a text roster"
            OutputAid.OutputTextRoster cboRoster.ListIndex, cboDate.Text, DelimNames, vbCrLf
            On Error GoTo 0
            
        Else 'RTF

            On Error GoTo TemplateError

            TempFileNum = FreeFile
            OutputAid.FileLoc = TempFileNum
            Open App.Path & "\~GVTempRoster.rtf" For Output As #TempFileNum

            TemplateNum = FreeFile
            Select Case cboRoster.ListIndex
                Case roAttendance: TemplateName = AttendanceTemplateName
                Case roCharacter: TemplateName = CharacterRosterTemplateName
                Case roPlayer: TemplateName = PlayerRosterTemplateName
            End Select
            Open App.Path & "\" & TemplateName For Input As #TemplateNum

            OutputAid.FilterRTFRoster TemplateNum, cboRoster.ListIndex, cboDate.Text, _
                    DelimNames, vbCrLf

            Close #TemplateNum
            Close #TempFileNum

            rtfEasel.Text = ""
            rtfEasel.LoadFile App.Path & "\~GVTempRoster.rtf"
            Kill App.Path & "\~GVTempRoster.rtf"

            rtfEasel.SelStart = Len(rtfEasel.Text)
            rtfEasel.SelLength = 0
            rtfEasel.SelText = kwPageBreak

            On Error GoTo PrintError

            rtfEasel.SelStart = 0
            Do

                If DocPrinting Then
                    PrinterAct = paNEWPAGE
                    Printer.NewPage
                End If
                DocPrinting = True

                rtfEasel.SelLength = InStr(rtfEasel.SelStart + 1, rtfEasel.Text, _
                        kwPageBreak, vbTextCompare) - rtfEasel.SelStart - 1

                PrinterAct = paPRINTRTF
                PrintRTF rtfEasel, 1440, 1440, 1440, 1440
                
                rtfEasel.SelStart = rtfEasel.SelStart + rtfEasel.SelLength + _
                        Len(kwPageBreak)

            Loop Until rtfEasel.SelStart = Len(rtfEasel.Text)

        End If

        Percent = Copy / Val(txtCopies)
        lblProgress = CStr(Int(Percent * 100)) & "%"
        lblMeter.Width = Percent * lblProgress.Width

        If DocPrinting Then
            PrinterAct = paENDDOC
            Printer.EndDoc
        End If
        
    Next Copy

    SaveSetting App.Title, "Output", "RTF Format", optRTF
    
    GoTo PrintFinish

PrintError:
    Screen.MousePointer = vbDefault
    If MsgBox("When Grapevine tried to" & PrinterAct & ", this error returned:" & _
            vbCrLf & Err.Description & vbCrLf & vbCrLf & "Attempt to continue?", _
            vbExclamation Or vbYesNo, "Printer Error") = vbYes Then
        Screen.MousePointer = vbHourglass
        Resume
    Else
        Resume PrintFinish
    End If

TemplateError:
    Screen.MousePointer = vbDefault
    MsgBox "Error writing temp RTF file: " & Err.Description, vbCritical + vbOKOnly, _
            "Temp file error"
    Close
    Resume PrintFinish

PrintFinish:

    PrinterAct = paDONE
    lblMeter.Visible = False
    lblProgress.Visible = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrintSetup_Click()
'
' Name:         cmdPrintSetup_Click
' Description:  Display the system's Print Setup window.
'
    
    Dim Device As Printer
    
    On Error Resume Next
    
    With cmnDialog
        .DialogTitle = "Printer Setup"
        .Flags = cdlPDPrintSetup + cdlPDReturnDC
        .ShowPrinter
    End With

    If Err.Number = 0 Then
        For Each Device In Printers
            If Device.hdc = cmnDialog.hdc Then Set Printer = Device
        Next Device
    End If
    
End Sub

Private Sub cmdSelectActive_Click()
'
' Name:         cmdSelectActive_Click
' Description:  Select all active characters.
'

    Dim Entry As Integer
    
    For Entry = 0 To lstMembers.ListCount - 1
        lstMembers.Selected(Entry) = (lstMembers.ItemData(Entry) = Active)
    Next Entry

End Sub

Private Sub cmdSelectAll_Click()
'
' Name:         cmdSelectAll_Click
' Description:  Select all players/characters.
'

    SelectMembers True

End Sub

Private Sub cmdClearSelect_Click()
'
' Name:         cmdClearSelect_Click
' Description:  Deselect all characters.
'

    SelectMembers False

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize all controls for the initial roster.
'

    Dim Entry As Integer
    Dim LongDate As Long

    cboDate.AddItem Format(Now, "mmmm d, yyyy")
    
    With Game.Calendar
        .Last
        Do Until .Off
            cboDate.AddItem Format(.GetGameDate, "mmmm d, yyyy")
            .MovePrevious
        Loop
    End With

    cboDate.Text = Format(Now, "mmmm d, yyyy")

    cboRoster.ListIndex = InitRoster
    
    optRTF = GetSetting(App.Title, "Output", "RTF Format", True)
    optPlainText = Not optRTF

End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
'
' Name:         txtCopies_KeyPress
' Description:  Ensure the number of copies is sane.
'

    If (Len(txtCopies) = 2 And KeyAscii <> vbKeyBack) Or _
       Not (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then
       KeyAscii = 0
    End If

End Sub
