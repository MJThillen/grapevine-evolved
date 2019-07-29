VERSION 5.00
Begin VB.Form frmNewActionRumor 
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
      Picture         =   "frmNewActionRumor.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAction 
      Height          =   240
      Left            =   240
      Picture         =   "frmNewActionRumor.frx":058A
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
Attribute VB_Name = "frmNewActionRumor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NewDate As Date
Public NewItem As String

Private IsAction As Boolean
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
    IsAction = False
    NewDate = RumorDate

    imgRumor.Visible = True
    cboTitles.Clear
    cboTitles.Visible = True

    lblLabel.Caption = " &Select or type in a title for the new rumor:"
    Me.Caption = "Add New Rumor"
    
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
                .MoveNext
            End If
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
    
    IsAction = False
    FindDate = True
    NewItem = RumorTitle
    Set InUseSet = New StringSet
    
    imgRumor.Visible = True
    lstDates.Clear
    lstDates.Visible = True

    lblLabel.Caption = " &Select a date for the new rumor:"
    Me.Caption = "Add New Rumor"
    
    SelSession = 0
    
    Game.APREngine.MoveToFirstTitle RumorList, RumorTitle
    Do Until RumorList.Off
        InUseSet.Add CStr(RumorList.Item.RumorDate)
        Game.APREngine.MoveToNextTitle RumorList, RumorTitle
    Loop
    
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
    
    IsAction = True
    FindDate = False
    NewDate = ActDate
    Set InUseSet = New StringSet
    
    imgAction.Visible = True
    cboSearches.Clear
    cboSearches.Visible = True
    lstChars.Visible = True

    lblLabel.Caption = " &Select the character who performs the new action:"
    Me.Caption = "Add New Action"
    
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
    
    IsAction = True
    FindDate = True
    NewItem = ActChar
    Set InUseSet = New StringSet
    
    imgAction.Visible = True
    lstDates.Clear
    lstDates.Visible = True
    
    lblLabel.Caption = " &Select a date for the character's new action:"
    Me.Caption = "Add New Action"
    
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
    
    UpdateCaption
    
    Me.Show vbModal, mdiMain

End Function

Private Sub UpdateCaption()
'
' Name:         UpdateCaption
' Description:  Update the caption of the txtPair control.
'

    txtPair.Text = IIf(IsAction, "Action:", "Rumor:") & vbCrLf & vbCrLf
    If FindDate Then
        txtPair.Text = txtPair.Text & lstDates.Text & vbCrLf & NewItem
    Else
        txtPair.Text = txtPair.Text & Format(NewDate, "mmmm d, yyyy") & vbCrLf & _
                        IIf(IsAction, lstChars.Text, cboTitles.Text)
    End If
    
End Sub

Private Sub cboSearches_Click()
'
' Name:         cboSearches_Click
' Description:  Change the list of characters.
'

    Dim Results As LinkedList
    Dim Values As LinkedList
    Dim Source As LinkedList
    
    Set Results = New LinkedList
    Set Values = New LinkedList
    lstChars.Clear
    
    Game.QueryEngine.QueryList.MoveTo cboSearches.Text
    If Game.QueryEngine.QueryList.Off Then
        Set Source = CharacterList
    Else
        Game.QueryEngine.MakeQuery Game.QueryEngine.QueryList.Item, Results, Values
        Set Source = Results
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

    Set Results = Nothing
    Set Values = Nothing

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

    Dim Source As LinkedList
    Dim NewName As String
    
    If FindDate Then
        If lstDates.ListIndex > -1 Then
            NewDate = CDate(lstDates.Text)
        Else
            NewDate = 0
        End If
    Else
        If IsAction Then
            NewItem = lstChars.Text
        Else
            NewItem = Trim(cboTitles.Text)
        End If
    End If
    
    If Not (NewDate = 0 Or NewItem = "") Then

        If IsAction Then
            Set Source = Game.APREngine.ActionList
        Else
            Set Source = Game.APREngine.RumorList
        End If
        
        Game.APREngine.MoveToPair Source, NewDate, NewItem

        If Not Source.Off Then
        
            MsgBox IIf(IsAction, "An action", "A rumor") & " already exists for the Date/Name pair " _
                   & vbCrLf & "you have selected.  Please provide a different pair.", vbOKOnly, _
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

