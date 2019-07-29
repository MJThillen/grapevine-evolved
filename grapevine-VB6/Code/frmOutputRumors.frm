VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOutputRumors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print or Export Action / Plot / Rumor Reports"
   ClientHeight    =   6915
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   8910
   Icon            =   "frmOutputRumors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optPerCharacter 
      Caption         =   "One File Per Character"
      Height          =   255
      Left            =   6480
      TabIndex        =   31
      Top             =   3240
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton optOneFile 
      Caption         =   "One File"
      Height          =   255
      Left            =   5040
      TabIndex        =   30
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox chkPlots 
      Caption         =   "Report Plot Developments (Master Reports Only)"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   1320
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.CheckBox chkRumors 
      Caption         =   "Report R&umors"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkActions 
      Caption         =   "Report Actions"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin ComCtl2.UpDown updCopies 
      Height          =   285
      Left            =   5700
      TabIndex        =   13
      Top             =   1320
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   503
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtCopies"
      BuddyDispid     =   196614
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
      TabIndex        =   12
      Text            =   "1"
      Top             =   1320
      Width           =   540
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "Printer Set&up..."
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectActive 
      Caption         =   """Acti&ve"" Status"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrintMasterList 
      Caption         =   "Print &Master Report"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdExportRumors 
      Caption         =   "&Export Reports"
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdExportMasterList 
      Caption         =   "Export Master Report"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrintRumors 
      Caption         =   "&Print Reports"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&All Characters"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearSelect 
      Caption         =   "Clea&r Selections"
      Height          =   375
      Left            =   1500
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ListBox lstCharacters 
      Columns         =   2
      Height          =   2400
      Left            =   480
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   5280
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   4920
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   $"frmOutputRumors.frx":058A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   240
      TabIndex        =   32
      Top             =   6120
      Width           =   8535
   End
   Begin VB.Label lblLabels 
      Caption         =   "Plain Text Only -- RTF not ready yet"
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
      Index           =   13
      Left            =   5040
      TabIndex        =   28
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Format"
      Height          =   195
      Index           =   11
      Left            =   5040
      TabIndex        =   26
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Copies:"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   11
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0%"
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   4920
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
      TabIndex        =   17
      Top             =   2880
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Index           =   7
      Left            =   4920
      TabIndex        =   18
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Select the Characters:"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   2
      Top             =   1920
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
      TabIndex        =   9
      Top             =   840
      Width           =   405
   End
   Begin VB.Label lblLabels 
      Caption         =   "Choose a &Date:"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1110
   End
   Begin VB.Label lblLabels 
      Caption         =   "Reports"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   23
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   5655
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Index           =   4
      Left            =   4920
      TabIndex        =   10
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lblMeter 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   12
      Left            =   4920
      TabIndex        =   27
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmOutputRumors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Active = 1
Const Other = 0

Private Sub SelectCharacters(Maybe As Boolean)
'
' Name:         SelectCharacters
' Parameters:   Maybe       whether to select or deselect all characters
' Description:  Select or deselect all characters.
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

Private Sub cmdExportMasterList_Click()
'
' Name:         cmdExportMasterList_Click
' Description:  Prompt the user for a filename, then save a Master List of
'               Rumors and/or Influence Use to a file.
'

    Dim FileNum As Integer
    Dim TemplateNum As Integer
    Dim CurrDate As Date

    If Not IsDate(cboDate.Text) Then Exit Sub

    With cmnDialog
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .DefaultExt = "txt"
        .InitDir = GetSetting(App.Title, "Files", "ExportDir", CurDir)
        .DialogTitle = "Save Master List As..."
        .Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        .FileName = "Master Report." & .DefaultExt
    End With

    On Error Resume Next
    cmnDialog.ShowSave
    If Err <> 0 Then Exit Sub 'Canceled

    On Error GoTo ExportMasterError

    Screen.MousePointer = vbHourglass
    CurrDate = CDate(cboDate.Text)

    FileNum = FreeFile
    Open cmnDialog.FileName For Output As #FileNum

    OutputAid.Destination = goFile
    OutputAid.FileLoc = FileNum

    OutputAid.SetStandardPageWidth

    Game.APREngine.OutputMasterReport OutputAid, CurrDate, (chkActions.Value = vbChecked), _
                                      (chkPlots.Value = vbChecked), (chkRumors.Value = vbChecked)

    GoTo ExportMasterFinish

ExportMasterError:
    Screen.MousePointer = vbDefault
    MsgBox "Problem saving file: " & Err.Description, vbOKOnly, "File Error"
    Resume ExportMasterFinish

ExportMasterFinish:

    Screen.MousePointer = vbDefault
    Close #FileNum

End Sub

Private Sub cmdExportRumors_Click()
'
' Name:         cmdExportRumors_Click
' Description:  Save Rumor/Influence Use reports to a file.  Text format
'               produced one large file, RTF produces many separate file.
'               Prompt the user for the location in which to save.
'

    Dim SaveType As SaveDirectoryType
    Dim Overwrite As Boolean

    Dim FileNum As Integer

    Dim CurrDate As Date
    Dim DocWriting As Boolean
    Dim Percent As Single
    Dim EntriesCounted As Integer
    
    Dim SelSet As StringSet
    Dim Skip As Boolean
    Dim Found As Boolean
    
    If Not IsDate(cboDate.Text) Then Exit Sub
    CurrDate = CDate(cboDate.Text)
    
    Set SelSet = New StringSet

    SelSet.StoreListBox lstCharacters

    With cmnDialog
        .InitDir = GetSetting(App.Title, "Files", "ExportDir", CurDir)
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .DefaultExt = "txt"
        .DialogTitle = "Save Reports As..."
        .FileName = "Reports.txt"
        .Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
    End With

    If optPerCharacter.Value Then

        If SelSet.Count > 1 Then
            frmSaveDirectory.GetSaveDirectory "the Reports:"
            Overwrite = frmSaveDirectory.Overwrite
            SaveType = frmSaveDirectory.Value
            Unload frmSaveDirectory
            If SaveType = sdCancel Then GoTo ExportFinish
        Else
            Overwrite = False
            SaveType = sdIndividual
        End If
        
    Else
    
        If SelSet.Count = 1 Then
            SelSet.First
            cmnDialog.FileName = ConvertToFileName(SelSet.StrItem) & " Report.txt"
        Else
            cmnDialog.FileName = "Reports.txt"
        End If
        On Error Resume Next
        cmnDialog.ShowSave
        If Err <> 0 Then GoTo ExportFinish
        On Error GoTo 0
        SaveType = sdOK
        
    End If
    
    On Error GoTo ExportError

    Screen.MousePointer = vbHourglass
    lblProgress = "0%"
    lblProgress.Visible = True
    lblMeter.Width = 0
    lblMeter.Visible = True
    EntriesCounted = 0

    Game.APREngine.PrepareRumorOutput CurrDate
 
    DocWriting = False
    
    If optOneFile.Value Then
        FileNum = FreeFile
        Open cmnDialog.FileName For Output As #FileNum
        OutputAid.Destination = goFile
        OutputAid.FileLoc = FileNum
        OutputAid.SetStandardPageWidth
    End If
    
    CharacterList.First
    Do Until CharacterList.Off
    
        If SelSet.Has(CharacterList.Item.Name) Then
    
            Skip = False
            Found = False
            
            If optPerCharacter.Value Then
            
                cmnDialog.FileName = ConvertToFileName(CharacterList.Item.Name) & " Report.txt"
            
                If SaveType = sdOK Then
                    Found = Dir(cmnDialog.FileName) <> "" And Not Overwrite
                End If
                
                If SaveType = sdIndividual Or Found Then
                
                    On Error Resume Next
                    cmnDialog.ShowSave
                    If Not Err = 0 Then Skip = True
                    On Error GoTo 0
                    
                End If

                If Not Skip Then
                    FileNum = FreeFile
                    Open cmnDialog.FileName For Output As #FileNum
                    OutputAid.Destination = goFile
                    OutputAid.FileLoc = FileNum
                    OutputAid.SetStandardPageWidth
                End If

            Else
                If DocWriting Then Print #FileNum, ""
            End If
            
            If Not Skip Then
            
                DocWriting = True

                Game.APREngine.OutputCharacterReport OutputAid, CharacterList.Item.Name, CurrDate, _
                               (chkActions.Value = vbChecked), (chkRumors.Value = vbChecked)

                If optPerCharacter.Value Then Close #FileNum

            End If

            EntriesCounted = EntriesCounted + 1
            Percent = EntriesCounted / SelSet.Count
            lblProgress = CStr(Int(Percent * 100)) & "%"
            lblMeter.Width = Percent * lblProgress.Width

        End If

        CharacterList.MoveNext
    Loop

    GoTo ExportFinish

ExportError:
    Screen.MousePointer = vbDefault
    MsgBox "Problem saving file: " & Err.Description, vbOKOnly, "File Error"
    Resume ExportFinish

ExportFinish:

    If DocWriting Then Close #FileNum
    Set SelSet = Nothing
    Screen.MousePointer = vbDefault
    lblMeter.Visible = False
    lblProgress.Visible = False

End Sub

Private Sub cmdPrintMasterList_Click()
'
' Name:         cmdPrintMasterList_Click
' Description:  Print a Master List of Rumors/Influence Use.
'

    Dim CurrDate As Date
    Dim DocPrinting As Boolean
    Dim Copy As Integer
    Dim Cancel As Boolean
    
    Dim FileNum As Integer
    Dim TemplateNum As Integer
    Dim TemplateName As String
    Dim TempFileNum As Integer
    Dim RTFPageLoc As Integer
    
    txtCopies = CStr(Val(txtCopies))
    If Val(txtCopies) = 0 Then Exit Sub
    If Not IsDate(cboDate.Text) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    DocPrinting = False
    CurrDate = CDate(cboDate.Text)
    
    On Error GoTo PrintMasterError
    
    PrinterAct = paSTART
    
    PrinterAct = paFONT
    Printer.Font.Name = "Courier New"
    Printer.Font.Size = 12
    PrinterAct = paWIDTH
    OutputAid.Destination = goPrinter
    OutputAid.SetStandardPageWidth Cancel
    If Cancel Then Exit Sub
    
    For Copy = 1 To Val(txtCopies)
        
        If DocPrinting Then
            PrinterAct = paNEWPAGE
            Printer.NewPage
        End If
        
        DocPrinting = True
        
        Game.APREngine.OutputMasterReport OutputAid, CurrDate, (chkActions.Value = vbChecked), _
                                          (chkPlots.Value = vbChecked), (chkRumors.Value = vbChecked)
        
    Next Copy

    If DocPrinting Then
        PrinterAct = paENDDOC
        Printer.EndDoc
    End If

    GoTo PrintMasterFinish

PrintMasterError:
    Screen.MousePointer = vbDefault
    If MsgBox("When Grapevine tried to" & PrinterAct & ", this error returned:" & _
            vbCrLf & Err.Description & vbCrLf & vbCrLf & "Attempt to continue?", _
            vbExclamation Or vbYesNo, "Printer Error") = vbYes Then
        Screen.MousePointer = vbHourglass
        Resume
    Else
        Resume PrintMasterFinish
    End If

TemplateMasterError:
    Screen.MousePointer = vbDefault
    MsgBox "Error working with template RTF file: " & Err.Description, vbCritical + vbOKOnly, _
            "Template file error"
    Close
    Resume PrintMasterFinish

PrintMasterFinish:
    
    PrinterAct = paDONE
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdPrintRumors_Click()
'
' Name:         cmdPrintRumors_Click
' Description:  Print Rumor/Influence Reports for characters.
'

    Dim Entry As Integer
    Dim CurrDate As Date
    Dim Header As String
    Dim DocPrinting As Boolean
    Dim Cancel As Boolean
    Dim Percent As Single
    Dim PresentCopy As Integer
    Dim Total As Integer
    Dim Copy As Integer

    Dim SelSet As StringSet

    Set SelSet = New StringSet

    SelSet.StoreListBox lstCharacters

    txtCopies = CStr(Val(txtCopies))
    Total = SelSet.Count * Val(txtCopies)
    
    If Total < 1 Or Not IsDate(cboDate.Text) Then GoTo PrintFinish

    CurrDate = CDate(cboDate.Text)
    PresentCopy = 0

    Screen.MousePointer = vbHourglass
    lblProgress = "0%"
    lblProgress.Visible = True
    lblMeter.Width = 0
    lblMeter.Visible = True

    On Error GoTo PrintError

    PrinterAct = paSTART

    PrinterAct = paFONT
    Printer.Font.Name = "Courier New"
    Printer.Font.Size = 12
    PrinterAct = paWIDTH
    OutputAid.SetStandardPageWidth Cancel
    If Cancel Then Exit Sub
    OutputAid.Destination = goPrinter
    OutputAid.FileLoc = -1
    Game.APREngine.PrepareRumorOutput CurrDate
    
    For Copy = 1 To Val(txtCopies)

        DocPrinting = False
        
        CharacterList.First
        Do Until CharacterList.Off
        
            If SelSet.Has(CharacterList.Item.Name) Then

                If DocPrinting Then
                    PrinterAct = paSTART
                    Printer.NewPage
                End If
                DocPrinting = True

                Game.APREngine.OutputCharacterReport OutputAid, CharacterList.Item.Name, CurrDate, _
                               (chkActions.Value = vbChecked), (chkRumors.Value = vbChecked)
                
                PresentCopy = PresentCopy + 1
                Percent = PresentCopy / Total
                lblProgress = CStr(Int(Percent * 100)) & "%"
                lblMeter.Width = Percent * lblProgress.Width

            End If

            CharacterList.MoveNext
        Loop

        If DocPrinting Then
            PrinterAct = paENDDOC
            Printer.EndDoc
        End If

    Next Copy

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

PrintFinish:

    Set SelSet = Nothing
    PrinterAct = paDONE
    Screen.MousePointer = vbDefault
    lblMeter.Visible = False
    lblProgress.Visible = False

End Sub

Private Sub cmdPrintSetup_Click()
'
' Name:         cmdPrintSetup_Click
' Description:  Display the system's Printer Setup dialog.
'
    
    On Error Resume Next
        
    With cmnDialog
        .DialogTitle = "Printer Setup"
        .Flags = cdlPDPrintSetup
        .ShowPrinter
    End With

End Sub

Private Sub cmdSelectActive_Click()
'
' Name:         cmdSelectActive_Click
' Description:  Select all Active characters.
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
' Description:  Populate all controls with the user's preferred options and with
'               game dates and characters.
'

    Dim Entry As Integer
    Dim NearDate As Date
    
    With Game.Calendar
        .MoveToCloseGame
        If Not .Off Then NearDate = .GetGameDate
        .First
        Do Until .Off
            cboDate.AddItem Format(.GetGameDate, "mmmm d, yyyy")
            If .GetGameDate = NearDate Then cboDate.ListIndex = cboDate.NewIndex
            .MoveNext
        Loop
    End With
    
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
        
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
'
' Name:         txtCopies_KeyPress
' Description:  Ensure that the number of copies entered is sane.
'

    If (Len(txtCopies) = 2 And KeyAscii <> vbKeyBack) Or _
       Not (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then
       KeyAscii = 0
    End If

End Sub
