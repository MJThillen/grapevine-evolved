VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOutputCards 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print or Export Item Cards, Rotes and Locations"
   ClientHeight    =   6165
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   8910
   Icon            =   "frmOutputCards.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstList 
      Columns         =   2
      Height          =   3570
      Left            =   480
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CommandButton cmdSelectActive 
      Caption         =   "Select Acti&ve"
      Height          =   375
      Left            =   1500
      TabIndex        =   22
      Top             =   5280
      Width           =   1935
   End
   Begin VB.ComboBox cboSelect 
      Height          =   315
      ItemData        =   "frmOutputCards.frx":058A
      Left            =   480
      List            =   "frmOutputCards.frx":05A0
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   720
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox rtfEasel 
      Height          =   495
      Left            =   5400
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmOutputCards.frx":0679
   End
   Begin MSComCtl2.UpDown updCopies 
      Height          =   285
      Left            =   5700
      TabIndex        =   6
      Top             =   1320
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   503
      _Version        =   393216
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
      TabIndex        =   5
      Text            =   "1"
      Top             =   1320
      Width           =   540
   End
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "Printer Set&up..."
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdExportSheets 
      Caption         =   "&Export Cards"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrintSheets 
      Caption         =   "&Print Cards"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearSelect 
      Caption         =   "Select &None"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
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
      Index           =   9
      Left            =   5040
      TabIndex        =   25
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   $"frmOutputCards.frx":06FB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   8
      Left            =   4920
      TabIndex        =   24
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Format"
      Height          =   195
      Index           =   11
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   12
      Left            =   4920
      TabIndex        =   19
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Copies:"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   4
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
      TabIndex        =   17
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
      Index           =   2
      Left            =   5040
      TabIndex        =   16
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   7
      Left            =   4920
      TabIndex        =   15
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Select the following to output:"
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
      TabIndex        =   14
      Top             =   840
      Width           =   405
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Items, Rotes, Locations"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   5655
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   4
      Left            =   4920
      TabIndex        =   13
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lblMeter 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "frmOutputCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Active = 1
Const Other = 0

Dim Selection As Integer

Private Function CountSelected()
'
' Name:         CountSelected
' Description:  Return the number of selected items.
' Returns:      the number of selected items.
'

    Dim Entry As Integer
    Dim Total As Integer

    For Entry = 0 To lstList.ListCount - 1
        If lstList.Selected(Entry) Then Total = Total + 1
    Next Entry
    
    CountSelected = Total

End Function

Private Sub SelectAll(Maybe As Boolean)
'
' Name:         SelectAll
' Parameters:   Maybe       whether to select or deselect all items
' Description:  Select or deselect all items.
'

    Dim Entry As Integer
    
    For Entry = 0 To lstList.ListCount - 1
        lstList.Selected(Entry) = Maybe
    Next Entry

End Sub

Private Sub cboSelect_Click()
'
' Name:         cboSelect_Click
' Description:  Populate the list appropriately.
'

    If Not cboSelect.ListIndex = Selection Then
    
        Selection = cboSelect.ListIndex
    
        lstList.Clear
        
        Select Case Selection
            Case 0, 2
                CharacterList.First
                Do Until CharacterList.Off
                    lstList.AddItem CharacterList.Item.Name
                    If CharacterList.Item.Status = ActiveStatus Then
                        lstList.ItemData(lstList.NewIndex) = Active
                    Else
                        lstList.ItemData(lstList.NewIndex) = Other
                    End If
                    CharacterList.MoveNext
                Loop
                cmdSelectActive.Visible = True
            Case 1
                CharacterList.First
                Do Until CharacterList.Off
                    If CharacterList.Item.RaceCode = gvracemage Then
                        lstList.AddItem CharacterList.Item.Name
                        If CharacterList.Item.Status = ActiveStatus Then
                            lstList.ItemData(lstList.NewIndex) = Active
                        Else
                            lstList.ItemData(lstList.NewIndex) = Other
                        End If
                    End If
                    CharacterList.MoveNext
                Loop
                cmdSelectActive.Visible = True
            Case 3
                ItemList.First
                Do Until ItemList.Off
                    lstList.AddItem ItemList.Item.Name
                    ItemList.MoveNext
                Loop
                cmdSelectActive.Visible = False
            Case 4
                RoteList.First
                Do Until RoteList.Off
                    lstList.AddItem RoteList.Item.Name
                    RoteList.MoveNext
                Loop
                cmdSelectActive.Visible = False
            Case 5
                LocationList.First
                Do Until LocationList.Off
                    lstList.AddItem LocationList.Item.Name
                    LocationList.MoveNext
                Loop
                cmdSelectActive.Visible = False
        End Select
    
    End If

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
    
    Dim DocWriting As Boolean
    Dim Percent As Single
    Dim TotalSelected As Integer
    Dim EntriesCounted As Integer
    
    Dim MainList As LinkedList
    Dim TraitList As LinkedTraitList
    Dim SubList As LinkedList
    Dim Header As String
    Dim SelSet As StringSet
    
    TotalSelected = CountSelected()
    If TotalSelected = 0 Then Exit Sub

    With cmnDialog
        .InitDir = GetSetting(App.Title, "Files", "DefaultDir", CurDir)
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .DefaultExt = "txt"
        .DialogTitle = "Save Card(s) As..."
        .FileName = "Cards.txt"
        .Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        On Error Resume Next
        .ShowSave
        If Err <> 0 Then Exit Sub 'Canceled
    End With
    SaveSetting App.Title, "Files", "DefaultDir", CurDir
    
    On Error GoTo ExportError
    
    Screen.MousePointer = vbHourglass
    lblProgress = "0%"
    lblProgress.Visible = True
    lblMeter.Width = 0
    lblMeter.Visible = True
    EntriesCounted = 0
    DocWriting = False
    
    Select Case Selection
        Case 0
            Set MainList = CharacterList
            Set SubList = ItemList
            Header = "Item Cards"
        Case 1
            Set MainList = CharacterList
            Set SubList = RoteList
            Header = "Rotes"
        Case 2
            Set MainList = CharacterList
            Set SubList = LocationList
            Header = "Favorite Locations"
        Case 3
            Set MainList = ItemList
        Case 4
            Set MainList = RoteList
        Case 5
            Set MainList = LocationList
    End Select
    
    Set SelSet = New StringSet
    SelSet.StoreListBox lstList
    MainList.First
    
    Do Until MainList.Off
        
        If SelSet.Has(MainList.Item.Name) Then
                
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
            
            If Selection < 3 Then
                OutputAid.Output MainList.Item.Name, Header
                Select Case Selection
                    Case 0: Set TraitList = MainList.Item.EquipmentList
                    Case 1: Set TraitList = MainList.Item.RoteList
                    Case 2: Set TraitList = MainList.Item.HangoutList
                End Select
                TraitList.First
                Do Until TraitList.Off
                    SubList.MoveTo TraitList.Trait.Name
                    If Not SubList.Off Then
                        SubList.Item.OutputTextSheet True, OutputAid
                    End If
                    TraitList.MoveNext
                Loop
            Else
                MainList.Item.OutputTextSheet True, OutputAid
            End If
            
            EntriesCounted = EntriesCounted + 1
            Percent = EntriesCounted / TotalSelected
            lblProgress = CStr(Int(Percent * 100)) & "%"
            lblMeter.Width = Percent * lblProgress.Width

        End If

        MainList.MoveNext

    Loop

    GoTo ExportFinish

ExportError:
    MsgBox "Problem saving file: " & Err.Description, vbOKOnly, "File Error"
    Resume ExportFinish

ExportFinish:

    If DocWriting Then Close #FileNum
    lblMeter.Visible = False
    lblProgress.Visible = False
    Set SelSet = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrintSheets_Click()
'
' Name:         cmdPrintSheets_Click
' Description:  Print the selected character sheets in the selected format.
'
        
    Dim DocPrinting As Boolean
    Dim Cancel As Boolean
    Dim Percent As Single
    Dim PresentCopy As Integer
    Dim Total As Integer
    Dim Copy As Integer
    
    Dim MainList As LinkedList
    Dim TraitList As LinkedTraitList
    Dim SubList As LinkedList
    Dim Header As String
    Dim SelSet As StringSet

    txtCopies = CStr(Val(txtCopies))
    Total = CountSelected() * Val(txtCopies)
    If Total < 1 Then Exit Sub
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
    OutputAid.SetStandardPageWidth Cancel
    If Cancel Then GoTo PrintFinish
    
    OutputAid.Destination = goPrinter
    
    Select Case Selection
        Case 0
            Set MainList = CharacterList
            Set SubList = ItemList
            Header = "Item Cards"
        Case 1
            Set MainList = CharacterList
            Set SubList = RoteList
            Header = "Rotes"
        Case 2
            Set MainList = CharacterList
            Set SubList = LocationList
            Header = "Favorite Locations"
        Case 3
            Set MainList = ItemList
        Case 4
            Set MainList = RoteList
        Case 5
            Set MainList = LocationList
    End Select
    
    Set SelSet = New StringSet
    SelSet.StoreListBox lstList
    
    For Copy = 1 To Val(txtCopies)
    
        DocPrinting = False
        MainList.First
        
        Do Until MainList.Off
        
            If SelSet.Has(MainList.Item.Name) Then
                                
                If Selection < 3 Then
                    PrinterAct = paNEWPAGE
                    If DocPrinting Then Printer.NewPage
                    DocPrinting = True
                    OutputAid.Output MainList.Item.Name, Header
                    Select Case Selection
                        Case 0: Set TraitList = MainList.Item.EquipmentList
                        Case 1: Set TraitList = MainList.Item.RoteList
                        Case 2: Set TraitList = MainList.Item.HangoutList
                    End Select
                    TraitList.First
                    Do Until TraitList.Off
                        SubList.MoveTo TraitList.Trait.Name
                        If Not SubList.Off Then
                            SubList.Item.OutputTextSheet True, OutputAid
                        End If
                        TraitList.MoveNext
                    Loop
                Else
                    DocPrinting = True
                    PrinterAct = "send text character sheet to printer"
                    MainList.Item.OutputTextSheet True, OutputAid
                End If
                
                PresentCopy = PresentCopy + 1
                Percent = PresentCopy / Total
                lblProgress = CStr(Int(Percent * 100)) & "%"
                lblMeter.Width = Percent * lblProgress.Width
    
            End If
    
            MainList.MoveNext
        
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
    Set SelSet = Nothing
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
    
    For Entry = 0 To lstList.ListCount - 1
        lstList.Selected(Entry) = (lstList.ItemData(Entry) = Active)
    Next Entry

End Sub

Private Sub cmdSelectAll_Click()
'
' Name:         cmdSelectAll_Click
' Description:  Select all characters.
'

    SelectAll True

End Sub

Private Sub cmdClearSelect_Click()
'
' Name:         cmdClearSelect_Click
' Description:  Deselect all characters.
'

    SelectAll False

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the list of characters and the preferred options.
'

    Selection = 99
    cboSelect.ListIndex = 0

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
