VERSION 5.00
Begin VB.Form frmGetAPRInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPair 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   1335
      Left            =   2520
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox lstDates 
      Height          =   2055
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cboTitles 
      Height          =   2055
      IntegralHeight  =   0   'False
      Left            =   240
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox lstChars 
      Height          =   1620
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cboSearches 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2340
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image imgPlot 
      Height          =   240
      Left            =   240
      Picture         =   "frmGetAPRInfo.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDate 
      Caption         =   "Previous Game:"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblDate 
      Caption         =   "Next Game:"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image imgRumor 
      Height          =   240
      Left            =   240
      Picture         =   "frmGetAPRInfo.frx":058A
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAction 
      Height          =   240
      Left            =   240
      Picture         =   "frmGetAPRInfo.frx":0B14
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblNext 
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lblPrevious 
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label lblLabel 
      Caption         =   " &Add new rumor for the following date and the given title:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmGetAPRInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NewDate As Date
Public NewItem As String

Private Target As APRType
Private FindDate As Boolean
Private InUseSet As StringSet

Public Function GetNewRumorTitle(RumorDate As Date)
'
' Name:         GetNewRumorTitle
' Parameters:   RumorDate       Date for which to create this rumor
' Description:  Determine a new rumor to be created with the given date and selected title.
'

    Dim TitleSet As StringSet
    Dim Title As String
    
    Set InUseSet = New StringSet
    Set TitleSet = New StringSet
    
    FindDate = False
    Target = aprRumor
    NewDate = RumorDate

    imgRumor.Visible = True
    cboTitles.Clear
    cboTitles.Visible = True
    
    lblLabel.Caption = " &Select or type in a title for the rumor:"
    Me.Caption = "Add Rumor"
    
    With Game.APREngine.RumorList
        .First
        Do Until .Off
            If Not .Item.Category = rtInfluence Then
                Title = .Item.Title
                If .Item.RumorDate = RumorDate Then
                    InUseSet.Add Title
                    TitleSet.Remove Title
                Else
                    If Not InUseSet.Has(Title) Then TitleSet.Add Title
                End If
            End If
            .MoveNext
        Loop
    End With

    TitleSet.First
    Do Until TitleSet.Off
        cboTitles.AddItem TitleSet.StrItem
        TitleSet.MoveNext
    Loop
    
    If cboTitles.ListCount > 0 Then
        cboTitles.ListIndex = 0
    End If

    UpdateCaption

    Set TitleSet = Nothing
    
    Me.Show vbModal, mdiMain

End Function

Public Function GetNewRumorDate(RumorTitle As String)
'
' Name:         GetNewRumorDate
' Parameters:   RumorTitle      the title for which to create this rumor
' Description:  Determine a new rumor to be created with the given title and selected date.
'

    Dim D As Date
    Dim SelSession As Date
    
    Target = aprRumor
    FindDate = True
    NewItem = RumorTitle
    Set InUseSet = New StringSet
    
    imgRumor.Visible = True
    lstDates.Clear
    lstDates.Visible = True

    lblLabel.Caption = " &Select a date for the rumor:"
    Me.Caption = "Add Rumor"
    
    SelSession = 0
    
    Game.APREngine.MoveToFirstTitle RumorList, RumorTitle
    Do Until RumorList.Off
        InUseSet.Add CStr(RumorList.Item.RumorDate)
        Game.APREngine.MoveToNextTitle RumorList, RumorTitle
    Loop
    
    With Game.Calendar
    
        .MoveToCloseGame
        If Not .Off Then SelSession = .GetGameDate
    
        .Last
        Do Until .Off
            D = .GetGameDate
            If Not InUseSet.Has(CStr(D)) Then
                lstDates.AddItem Format(D, "mmmm d, yyyy")
                If D = SelSession Then lstDates.ListIndex = lstDates.NewIndex
            End If
            .MovePrevious
        Loop
    
    End With
    
    If lstDates.ListIndex = -1 And lstDates.ListCount > 0 Then lstDates.ListIndex = 0
    UpdateCaption
    
    Me.Show vbModal, mdiMain

End Function

Public Function GetNewPlotTitle(PlotDate As Date)
'
' Name:         GetNewPlotTitle
' Parameters:   PlotDate       Date for which to create this Plot
' Description:  Determine a new Plot to be created with the given date and selected title.
'

    Dim Title As String
    Dim Plot As PlotClass
    
    FindDate = False
    Target = aprPlot
    NewDate = PlotDate

    imgPlot.Visible = True
    lstDates.Clear
    lstDates.Visible = True
    
    lblLabel.Caption = " &Select a plot to develop:"
    Me.Caption = "Add Plot Development"
    
    With Game.APREngine.PlotList
        .First
        Do Until .Off
            Set Plot = .Item
            If Plot.GetStatus(NewDate) = psActive Then
                Plot.MoveTo NewDate
                If Plot.Off Then
                    lstDates.AddItem Plot.Name
                End If
            End If
            .MoveNext
        Loop
    End With

    If lstDates.ListCount > 0 Then
        lstDates.ListIndex = 0
    End If

    UpdateCaption

    Me.Show vbModal, mdiMain

End Function

Public Function GetNewPlotDate(PlotTitle As String)
'
' Name:         GetNewPlotDate
' Parameters:   PlotTitle      the title for which to create this Plot
' Description:  Determine a new Plot to be created with the given title and selected date.
'

    Dim D As Date
    Dim SelSession As Date
    Dim Plot As PlotClass
    
    Target = aprPlot
    FindDate = True
    NewItem = PlotTitle
    Set InUseSet = New StringSet
    
    imgPlot.Visible = True
    lstDates.Clear
    lstDates.Visible = True

    lblLabel.Caption = " &Select a date for the plot development:"
    Me.Caption = "Add Plot"
    
    SelSession = 0
    
    PlotList.MoveTo NewItem
    If Not PlotList.Off Then
        Set Plot = PlotList.Item
        Plot.First
        Do Until Plot.Off
            InUseSet.Add CStr(Plot.PlotDev.DevDate)
            Plot.MoveNext
        Loop
        PlotList.MoveNext
    End If
    
    With Game.Calendar
    
        .MoveToCloseGame
        If Not .Off Then SelSession = .GetGameDate
    
        .Last
        Do Until .Off
            D = .GetGameDate
            If Not InUseSet.Has(CStr(D)) Then
                lstDates.AddItem Format(D, "mmmm d, yyyy")
                If D = SelSession Then lstDates.ListIndex = lstDates.NewIndex
            End If
            .MovePrevious
        Loop
    
    End With
    
    If lstDates.ListIndex = -1 And lstDates.ListCount > 0 Then lstDates.ListIndex = 0
    UpdateCaption
    
    Me.Show vbModal, mdiMain

End Function

Public Function GetNewActionChar(ActDate As Date)
'
' Name:         GetNewAction
' Parameters:   ActDate     Date for which to create this Action
' Description:  Determine a new action to be created with the given date and selected character.
'

    Dim I As Integer
    
    Target = aprAction
    FindDate = False
    NewDate = ActDate
    Set InUseSet = New StringSet
    
    imgAction.Visible = True
    cboSearches.Clear
    cboSearches.Visible = True
    lstChars.Visible = True

    lblLabel.Caption = " &Select the character who performs the action:"
    Me.Caption = "Add Action"
    
    Game.APREngine.MoveToFirstDate ActionList, ActDate
    Do Until ActionList.Off
        InUseSet.Add ActionList.Item.CharName
        Game.APREngine.MoveToNextDate ActionList, ActDate
    Loop
    
    With Game.QueryEngine.QueryList
        .First
        Do Until .Off
            If .Item.Inventory = qiCharacters Then cboSearches.AddItem .Item.Name
            .MoveNext
        Loop
    End With

    For I = 0 To cboSearches.ListCount
        If cboSearches.List(I) = "Active Characters" Then
            cboSearches.ListIndex = I
            Exit For
        End If
    Next I

    If cboSearches.ListIndex = -1 Then Call cboSearches_Click

    UpdateCaption

    Me.Show vbModal, mdiMain

End Function

Public Function GetNewActionDate(ActChar As String)
'
' Name:         GetNewActionDate
' Parameters:   ActChar     The character for which to create the action
' Description:  Determine a new action to be created with the given character and selected date.
'

    Dim D As Date
    Dim SelSession As Date
    
    Target = aprAction
    FindDate = True
    NewItem = ActChar
    Set InUseSet = New StringSet
    
    imgAction.Visible = True
    lstDates.Clear
    lstDates.Visible = True
    
    lblLabel.Caption = " &Select a date for the character's action:"
    Me.Caption = "Add Action"
    
    Game.APREngine.MoveToFirstTitle ActionList, ActChar
    Do Until ActionList.Off
        InUseSet.Add CStr(ActionList.Item.ActDate)
        Game.APREngine.MoveToNextTitle ActionList, ActChar
    Loop
    
    SelSession = 0
    
    With Game.Calendar
    
        If .HasNextGame Then SelSession = .NextGameDate
        If SelSession = 0 And .HasPreviousGame Then SelSession = .PreviousGameDate
    
        .Last
        Do Until .Off
            D = .GetGameDate
            If Not InUseSet.Has(CStr(D)) Then
                lstDates.AddItem Format(D, "mmmm d, yyyy")
                If D = SelSession Then lstDates.ListIndex = lstDates.NewIndex
            End If
            .MovePrevious
        Loop
    
    End With
    
    If lstDates.ListIndex = -1 And lstDates.ListCount > 0 Then lstDates.ListIndex = 0
    UpdateCaption
    
    Me.Show vbModal, mdiMain

End Function

Private Sub UpdateCaption()
'
' Name:         UpdateCaption
' Description:  Update the caption of the txtPair control.
'

    Dim Cap As String
    Dim SDate As String
    
    Select Case Target
        Case aprAction: Cap = "Action:"
        Case aprPlot:   Cap = "Plot:"
        Case aprRumor:  Cap = "Rumor:"
    End Select
    
    Cap = Cap & vbCrLf & vbCrLf
    
    If FindDate Then
        txtPair.Text = Cap & lstDates.Text & vbCrLf & NewItem
    Else
        SDate = Format(NewDate, "mmmm d, yyyy")
        Select Case Target
            Case aprAction: txtPair.Text = Cap & SDate & vbCrLf & lstChars.Text
            Case aprPlot:   txtPair.Text = Cap & lstDates.Text & vbCrLf & SDate
            Case aprRumor:  txtPair.Text = Cap & SDate & vbCrLf & cboTitles.Text
        End Select
    End If
    
End Sub

Private Sub cboSearches_Click()
'
' Name:         cboSearches_Click
' Description:  Change the list of characters.
'

    Dim Source As LinkedList
    
    lstChars.Clear
    
    Game.QueryEngine.QueryList.MoveTo cboSearches.Text
    If Game.QueryEngine.QueryList.Off Then
        Set Source = CharacterList
    Else
        Game.QueryEngine.MakeQuery Game.QueryEngine.QueryList.Item
        Set Source = Game.QueryEngine.Results
    End If
    
    Source.First
    Do Until Source.Off
        If Not InUseSet.Has(Source.Item.Name) Then
            lstChars.AddItem Source.Item.Name
        End If
        Source.MoveNext
    Loop
    
    If lstChars.ListCount = 0 Then
        UpdateCaption
    Else
        lstChars.ListIndex = 0
    End If

End Sub

Private Sub cboTitles_Change()
'
' Name:         cboTitles_Change
' Description:  Update the txtPair.
'
    UpdateCaption

End Sub

Private Sub cboTitles_DblClick()
'
' Name:         cboTitles_DblClick
' Description:  Shortcut to the OK button.
'

    Call cmdOK_Click

End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Nullify the selected values and cancel the dialog.
'

    NewDate = 0
    NewItem = ""
    Set InUseSet = Nothing
    Me.Hide

End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Validate the user's choices and close the dialog.
'

    Dim NewName As String
    Dim Found As Boolean
    
    If FindDate Then
        If lstDates.ListIndex > -1 Then
            NewDate = CDate(lstDates.Text)
        Else
            NewDate = 0
        End If
    Else
        Select Case Target
            Case aprAction: NewItem = lstChars.Text
            Case aprPlot:   NewItem = lstDates.Text
            Case aprRumor:  NewItem = Trim(cboTitles.Text)
        End Select
    End If
    
    If Not (NewDate = 0 Or NewItem = "") Then

        If Target = aprPlot Then
            PlotList.MoveTo NewItem
            Found = Not PlotList.Off
            If Found Then
                PlotList.Item.MoveTo NewDate
                Found = Not PlotList.Item.Off
            End If
        Else
            If Target = aprAction Then
                Game.APREngine.MoveToPair ActionList, NewDate, NewItem
                Found = Not ActionList.Off
            Else
                Game.APREngine.MoveToPair RumorList, NewDate, NewItem
                Found = Not RumorList.Off
            End If
        End If

        If Found Then
        
            MsgBox "The Date/Name pair you have selected is in use." & vbCrLf & _
                   "Please provide a different pair.", vbOKOnly, _
                   "Date and Name in Use"
        
        Else
        
            Set InUseSet = Nothing
            Me.Hide
        
        End If

    End If

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Populate the game dates when the form loads.
'

    With Game.Calendar
    
        If .HasNextGame Then
            lblNext.Caption = Format(.NextGameDate, "mmmm d, yyyy")
        End If
        
        If .HasPreviousGame Then
            lblPrevious.Caption = Format(.PreviousGameDate, "mmmm d, yyyy")
        End If
        
    End With
    
End Sub

Private Sub lstChars_Click()
'
' Name:         lstChars_Click
' Description:  Update the txtPair.
'
    UpdateCaption

End Sub

Private Sub lstChars_DblClick()
'
' Name:         lstChars_DblClick
' Description:  Shortcut to the OK button.
'
    Call cmdOK_Click
    
End Sub

Private Sub lstDates_Click()
'
' Name:         lstDates_Click
' Description:  Update the txtPair.
'
    UpdateCaption

End Sub

Private Sub lstDates_DblClick()
'
' Name:         lstDates_DblClick
' Description:  Update the txtPair.
'
    Call cmdOK_Click

End Sub

