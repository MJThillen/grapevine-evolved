VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectLink 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Cause"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2580
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin MSComctlLib.TreeView tvwTree 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4048
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   354
      LabelEdit       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSelectLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ChoiceType As APRType
Public Item As String
Public Subitem As String
Public When As Date
Private Effect As Boolean

Public Sub SelectLink(ForDate As Date, IsEffect As Boolean)
'
' Name:         Selectlink
' Parameters:   ForDate         date for which this link is made
'               IsEffect        TRUE iff this is an effect, FALSE iff it's a cause
' Description:  Show a tree of actions and plots (and rumors if this is an effect link)
'               for the given date, from which the user can choose.
'
    
    Dim I As Integer
    
    When = ForDate
    Effect = IsEffect
    Me.Caption = "Select " & IIf(Effect, "Effect", "Cause")
    cboDate.Clear
    
    With Game.Calendar
        .First
        Do Until .Off
            cboDate.AddItem Format(.GetGameDate, "mmmm d, yyyy")
            If .GetGameDate = When Then I = cboDate.NewIndex
            .MoveNext
        Loop
    End With
    
    cboDate.ListIndex = I
    Me.Show vbModal, mdiMain
        
End Sub

Private Sub RefreshLinks()
'
' Name:         RefreshLinks
' Description:  Refresh the tree of links based on the given date.
'


    Dim TopNode As Node
    Dim TempNode As Node
    Dim TagNode As Node
    Dim Action As ActionClass
    Dim Plot As PlotClass
    Dim Rumor As RumorClass
    Dim Explain As String

    tvwTree.Nodes.Clear
    Explain = IIf(Effect, "Affect ", "Be caused by ")
    
    Set TopNode = tvwTree.Nodes.Add(Key:="Top", _
                  Text:=Format(When, "mmmm d") & IIf(Effect, " Effects", " Causes"))
    
    Set TempNode = tvwTree.Nodes.Add("Top", tvwChild, "Actions", Explain & "an Action")
    TempNode.Sorted = True
    TempNode.Expanded = False
    Set TempNode = tvwTree.Nodes.Add("Top", tvwChild, "Plots", Explain & "a Plot")
    TempNode.Sorted = True
    TempNode.Expanded = False
    
    If Effect Then
        Set TempNode = tvwTree.Nodes.Add("Top", tvwChild, "Rumors", Explain & "a Rumor")
        TempNode.Sorted = False
        TempNode.Expanded = False
    End If
    
    Game.APREngine.MoveToFirstDate ActionList, When
    Do Until ActionList.Off
        
        Set Action = ActionList.Item
        Set TempNode = tvwTree.Nodes.Add("Actions", tvwChild, , Action.CharName)
        
        If Action.Count > 1 Then
            Action.First
            Do Until Action.Off
                Set TagNode = tvwTree.Nodes.Add(TempNode.Index, tvwChild, , Action.SubAction.Name)
                TagNode.Tag = "a" & Action.CharName & vbCr & Action.SubAction.Name
                Action.MoveNext
            Loop
        Else
            TempNode.Tag = "a" & Action.CharName & vbCr & BasicSubactionName
        End If
        
        Game.APREngine.MoveToNextDate ActionList, When
        
    Loop
    
    PlotList.First
    Do Until PlotList.Off
        
        Set Plot = PlotList.Item
        
        Plot.MoveTo When
        If Not Plot.Off Then
            Set TempNode = tvwTree.Nodes.Add("Plots", tvwChild, , Plot.Name)
            TempNode.Tag = "p" & Plot.Name & vbCr
        End If
        
        PlotList.MoveNext
    
    Loop
    
    If Effect Then
    
        Game.APREngine.MoveToFirstDate RumorList, When
        Do Until RumorList.Off
            
            Set Rumor = RumorList.Item
            Set TempNode = tvwTree.Nodes.Add("Rumors", tvwChild, , Rumor.Title)
            
            If Rumor.Count > 1 Then
                Rumor.First
                Do Until Rumor.Off
                    Set TagNode = tvwTree.Nodes.Add(TempNode.Index, tvwChild, _
                                  , "Level " & CStr(Rumor.SubRumor.Level))
                    TagNode.Tag = "r" & Rumor.Title & vbCr & CStr(Rumor.SubRumor.Level)
                    Rumor.MoveNext
                Loop
            Else
                TempNode.Tag = "r" & Rumor.Title & vbCr & "0"
            End If
            
            Game.APREngine.MoveToNextDate RumorList, When
            
        Loop
    
    End If
    
    Set tvwTree.SelectedItem = TopNode
    
End Sub

Private Sub cboDate_Click()
'
' Name:         cboDate_Click
' Description:  Change the date and the display of links.
'

    When = CDate(cboDate.Text)
    RefreshLinks

End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdOK_Click
' Description:  Cancel the link selection.
'
    ChoiceType = aprNone
    Item = ""
    Subitem = ""
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Select the chosen link.
'
    
    If Not tvwTree.SelectedItem Is Nothing Then
    
        Dim NodeTag As String
        
        NodeTag = tvwTree.SelectedItem.Tag
        If NodeTag <> "" Then
        
            Dim Delim As Integer
        
            Select Case Left(NodeTag, 1)
                Case "a":   ChoiceType = aprAction
                Case "p":   ChoiceType = aprPlot
                Case "r":   ChoiceType = aprRumor
            End Select
            
            NodeTag = Mid(NodeTag, 2)
            Delim = InStr(NodeTag, vbCr)
            
            Item = Left(NodeTag, Delim - 1)
            Subitem = Mid(NodeTag, Delim + 1)
            
            Me.Hide
        
        End If
        
    End If

End Sub

Private Sub tvwTree_Click()
'
' Name:         tvwTree_Click
' Description:  Enable/Disable the OK button as needed.
'

    If Not tvwTree.SelectedItem Is Nothing Then
        cmdOK.Enabled = Not (tvwTree.SelectedItem.Tag = "")
    End If

End Sub

Private Sub tvwTree_DblClick()
'
' Name:         tvwTree_DblClick
' Description:  Shortcut to the OK button.
'

    Call cmdOK_Click

End Sub
