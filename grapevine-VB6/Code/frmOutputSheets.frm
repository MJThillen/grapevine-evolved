VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOutputSheets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print or Export Character Sheets"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   8910
   Icon            =   "frmOutputSheets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
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
      TextRTF         =   $"frmOutputSheets.frx":058A
   End
   Begin VB.OptionButton optPlainText 
      Caption         =   "Plain Text"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optRTF 
      Caption         =   "Rich Text Format (RTF)"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox chkHistory 
      Caption         =   "Include Experience &History"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   720
      Width           =   3375
   End
   Begin VB.CheckBox chkNotes 
      Caption         =   "Include &Notes"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin MSComCtl2.UpDown updCopies 
      Height          =   285
      Left            =   5700
      TabIndex        =   12
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
      TabIndex        =   11
      Text            =   "1"
      Top             =   2520
      Width           =   540
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "Printer Set&up..."
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectActive 
      Caption         =   """Acti&ve"" Status"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdExportSheets 
      Caption         =   "&Export Character Sheets"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrintSheets 
      Caption         =   "&Print Character Sheets"
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&All Characters"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearSelect 
      Caption         =   "&Clear Selections"
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ListBox lstCharacters 
      Columns         =   2
      Height          =   3570
      Left            =   480
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
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
      Caption         =   "&Format"
      Height          =   195
      Index           =   11
      Left            =   5040
      TabIndex        =   7
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
      Caption         =   "Options"
      Height          =   195
      Index           =   9
      Left            =   5040
      TabIndex        =   26
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sheets:"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   25
      Top             =   720
      Width           =   4020
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Copies:"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   10
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   3600
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   7
      Left            =   4920
      TabIndex        =   21
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Select the Characters for whom to Generate Character"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   0
      Top             =   480
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
      TabIndex        =   20
      Top             =   2040
      Width           =   405
   End
   Begin VB.Label lblLabels 
      Caption         =   "Character Sheets"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   5655
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   4
      Left            =   4920
      TabIndex        =   19
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblMeter 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Index           =   10
      Left            =   4920
      TabIndex        =   27
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmOutputSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Active = 1
Const Other = 0

Private Function CountSelected()
'
' Name:         CountSelected
' Description:  Return the number of selected items.
' Returns:      the number of selected items.
'

    Dim Entry As Integer
    Dim Total As Integer

    For Entry = 0 To lstCharacters.ListCount - 1
        If lstCharacters.Selected(Entry) Then Total = Total + 1
    Next Entry
    
    CountSelected = Total

End Function

Private Sub SelectCharacters(Maybe As Boolean)
'
' Name:         SelectCharacters
' Parameters:   Maybe       whether to select or deselect all items
' Description:  Select or deselect all items.
'

    Dim Entry As Integer
    
    For Entry = 0 To lstCharacters.ListCount - 1
        lstCharacters.Selected(Entry) = Maybe
    Next Entry

End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Dismiss this window.
'
    
    Unload Me

End Sub

Private Sub cmdExportSheets_Click()
'
' Name:         cmdExportSheets_Click
' Description:  Save the selected character sheets to file.  If they are in
'               plain text, save to one large file; for RTF, save in individual
'               files.  Prompt for the save location.
'

    Dim SaveType As SaveDirectoryType
    Dim Overwrite As Boolean
    Dim Continue As Boolean
    
    Dim FileNum As Integer
    Dim TemplateNum As Integer
    Dim TemplateName As String
    
    Dim Entry As Integer
    Dim DocWriting As Boolean
    Dim Percent As Single
    Dim TotalSelected As Integer
    Dim EntriesCounted As Integer
    
    TotalSelected = CountSelected
    If TotalSelected = 0 Then Exit Sub
    
    If optPlainText Then
    
        With cmnDialog
            .InitDir = GetSetting(App.Title, "Files", "DefaultDir", CurDir)
            .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .DefaultExt = "txt"
            .DialogTitle = "Save Character Sheet(s) As..."
            .FileName = "Sheets.txt"
            .Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
            On Error Resume Next
            .ShowSave
            If Err <> 0 Then Exit Sub 'Canceled
        End With
        SaveSetting App.Title, "Files", "DefaultDir", CurDir
        
    Else
    
        If TotalSelected > 1 Then
            frmSaveDirectory.GetSaveDirectory "the Character Sheets:"
            Overwrite = frmSaveDirectory.Overwrite
            SaveType = frmSaveDirectory.Value
            Unload frmSaveDirectory
            If SaveType = sdcancel Then Exit Sub
        End If
        
        With cmnDialog
            .Filter = "Rich Text Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
            .DefaultExt = "rtf"
            .DialogTitle = "Save Character Sheet As..."
            .Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        End With
    
    End If
    
    On Error GoTo ExportError
    
    Screen.MousePointer = vbHourglass
    lblProgress = "0%"
    lblProgress.Visible = True
    lblMeter.Width = 0
    lblMeter.Visible = True
    EntriesCounted = 0
    
    DocWriting = False
    For Entry = 0 To lstCharacters.ListCount - 1
        
        If lstCharacters.Selected(Entry) Then
            CharacterList.MoveTo lstCharacters.List(Entry)
            If Not CharacterList.Off Then
            
                If optPlainText Then
                    
                    If DocWriting Then
                        Print #FileNum, ""
                    Else
                        FileNum = FreeFile
                        Open cmnDialog.FileName For Output As #FileNum
                        OutputAid.Destination = goFile
                        OutputAid.FileLoc = FileNum
                        OutputAid.SetStandardPageWidth
                    End If
                    
                    DocWriting = True
                    CharacterList.Item.OutputTextSheet chkNotes, chkHistory, OutputAid
            
                Else 'optRTF
                    
                    Continue = True
                    
                    If SaveType = sdIndividual Or TotalSelected = 1 Then
                        With cmnDialog
                            .InitDir = GetSetting(App.Title, "Files", "DefaultDir", CurDir)
                            .FileName = ConvertToFileName(CharacterList.Item.Name) & ".rtf"
                            On Error GoTo ExportFinish
                            .ShowSave
                        End With
                        SaveSetting App.Title, "Files", "DefaultDir", CurDir
                    Else
                        cmnDialog.FileName = ConvertToFileName(CharacterList.Item.Name) & ".rtf"
                        If Not Overwrite And Dir(cmnDialog.FileName) <> "" Then
                            Select Case MsgBox("Overwrite file """ & CharacterList.Item.Name _
                                    & ".rtf"" ?", vbYesNoCancel + vbQuestion, "Overwrite")
                                Case vbNo
                                    Continue = False
                                Case vbCancel
                                    Exit For
                            End Select
                        End If
                    End If
                
                    If Continue Then
                
                        On Error GoTo ExportError
                    
                        FileNum = FreeFile
                        Open cmnDialog.FileName For Output As #FileNum
                        OutputAid.Destination = goFile
                        OutputAid.FileLoc = FileNum
                        
                        TemplateName = CharacterList.Item.GetTemplateSheet
                    
                        If TemplateName <> "" Then
                                TemplateNum = FreeFile
                                Open App.Path & "\" & TemplateName For Input As #TemplateNum
                                OutputAid.FilterRTFCharacter TemplateNum, CharacterList.Item, _
                                        chkNotes, chkHistory
                                Close #TemplateNum
                        End If
                        
                        Close #FileNum
                    
                    End If
                    
                End If
                
                EntriesCounted = EntriesCounted + 1
                Percent = EntriesCounted / TotalSelected
                lblProgress = CStr(Int(Percent * 100)) & "%"
                lblMeter.Width = Percent * lblProgress.Width
            
            End If
        End If
        
    Next Entry
    
    SaveSetting App.Title, "Output", "Notes", chkNotes
    SaveSetting App.Title, "Output", "History", chkHistory
    SaveSetting App.Title, "Output", "RTF Format", optRTF
    
    GoTo ExportFinish
    
ExportError:
    MsgBox "Problem saving file: " & Err.Description, vbOKOnly, "File Error"
    Resume ExportFinish

ExportFinish:
    
    If DocWriting Then Close #FileNum
    Screen.MousePointer = vbDefault
    lblMeter.Visible = False
    lblProgress.Visible = False
    
End Sub

Private Sub cmdPrintSheets_Click()
'
' Name:         cmdPrintSheets_Click
' Description:  Print the selected character sheets in the selected format.
'
        
    Dim Entry As Integer
    Dim DocPrinting As Boolean
    Dim Cancel As Boolean
    Dim Percent As Single
    Dim PresentCopy As Integer
    Dim Total As Integer
    Dim Copy As Integer
    
    Dim TemplateNum As Integer
    Dim TemplateName As String
    Dim TempFileNum As Integer
    Dim RTFPageLoc As Integer

    txtCopies = CStr(Val(txtCopies))
    Total = CountSelected * Val(txtCopies)
    If Total < 1 Then Exit Sub
    PresentCopy = 0
    
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
        OutputAid.SetStandardPageWidth Cancel
        If Cancel Then Exit Sub
    End If
    
    OutputAid.Destination = goPrinter
    
    For Copy = 1 To Val(txtCopies)
    
        DocPrinting = False
        For Entry = 0 To lstCharacters.ListCount - 1
        
            If lstCharacters.Selected(Entry) Then
                CharacterList.MoveTo lstCharacters.List(Entry)
                If Not CharacterList.Off Then
                
                    If optPlainText Then
                        
                        PrinterAct = paNEWPAGE
                        If DocPrinting Then Printer.NewPage
                        DocPrinting = True
                        PrinterAct = "send text character sheet to printer"
                        CharacterList.Item.OutputTextSheet _
                                chkNotes, chkHistory, OutputAid
                    
                    Else 'RTF
                    
                        On Error GoTo TemplateError
                        
                        TempFileNum = FreeFile
                        OutputAid.FileLoc = TempFileNum
                        TemplateName = CharacterList.Item.GetTemplateSheet
                    
                        Open App.Path & "\~GVTempSheet.rtf" For Output As #TempFileNum
                        
                        TemplateNum = FreeFile
                        Open App.Path & "\" & TemplateName For Input As #TemplateNum
                        
                        OutputAid.FilterRTFCharacter TemplateNum, CharacterList.Item, _
                                chkNotes, chkHistory
                        
                        Close #TemplateNum
                        Close #TempFileNum
                        
                        rtfEasel.Text = ""
                        rtfEasel.LoadFile App.Path & "\~GVTempSheet.rtf"
                        Kill App.Path & "\~GVTempSheet.rtf"
                        
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
                
                    PresentCopy = PresentCopy + 1
                    Percent = PresentCopy / Total
                    lblProgress = CStr(Int(Percent * 100)) & "%"
                    lblMeter.Width = Percent * lblProgress.Width
                
                End If
            End If
            
        Next Entry

        If DocPrinting Then
            PrinterAct = paENDDOC
            Printer.EndDoc
        End If
        
    Next Copy

    SaveSetting App.Title, "Output", "Notes", chkNotes
    SaveSetting App.Title, "Output", "History", chkHistory
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
    MsgBox "Error working with template RTF file: " & Err.Description, vbCritical + vbOKOnly, _
            "Template file error"
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
' Description:  Display the system's Print Setup dialog box.
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
    
    For Entry = 0 To lstCharacters.ListCount - 1
        lstCharacters.Selected(Entry) = (lstCharacters.ItemData(Entry) = Active)
    Next Entry

End Sub

Private Sub cmdSelectAll_Click()
'
' Name:         cmdSelectAll_Click
' Description:  Select all characters.
'

    SelectCharacters True

End Sub

Private Sub cmdClearSelect_Click()
'
' Name:         cmdClearSelect_Click
' Description:  Deselect all characters.
'

    SelectCharacters False

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the list of characters and the preferred options.
'

    Dim LongDate As Long
    Dim Entry As Integer

    CharacterList.First
    Do Until CharacterList.Off
        lstCharacters.AddItem CharacterList.Item.Name
        Select Case CharacterList.Item.Status
            Case ActiveStatus
                lstCharacters.ItemData(lstCharacters.NewIndex) = Active
            Case Else
                lstCharacters.ItemData(lstCharacters.NewIndex) = Other
        End Select
        CharacterList.MoveNext
    Loop

    chkNotes = GetSetting(App.Title, "Output", "Notes", vbUnchecked)
    chkHistory = GetSetting(App.Title, "Output", "History", vbUnchecked)
    optRTF = GetSetting(App.Title, "Output", "RTF Format", True)
    optPlainText = Not optRTF

End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
'
' Name:         txtCopies_KeyPress
' Description:  Ensure the number of copies entered is sane.
'

    If (Len(txtCopies) = 2 And KeyAscii <> vbKeyBack) Or _
       Not (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then
       KeyAscii = 0
    End If

End Sub
