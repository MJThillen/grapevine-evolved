VERSION 5.00
Begin VB.Form frmStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Statistics"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmStatistics.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9060
   Begin VB.PictureBox picView 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   360
      ScaleHeight     =   60.59
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   147.373
      TabIndex        =   23
      Top             =   2160
      Width           =   8415
      Begin VB.PictureBox picMatches 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   6240
         ScaleHeight     =   3585
         ScaleWidth      =   2025
         TabIndex        =   30
         Top             =   -120
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton cmdHide 
            Caption         =   "×"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1785
            TabIndex        =   36
            Top             =   105
            Width           =   255
         End
         Begin VB.ListBox lstMatches 
            Appearance      =   0  'Flat
            Height          =   2145
            IntegralHeight  =   0   'False
            ItemData        =   "frmStatistics.frx":058A
            Left            =   120
            List            =   "frmStatistics.frx":058C
            TabIndex        =   31
            Top             =   630
            Width           =   1755
         End
         Begin VB.CommandButton cmdShowCharacter 
            Caption         =   "Show Character"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   2880
            Width           =   1755
         End
         Begin VB.Label lblMatchCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   420
            Width           =   1755
         End
         Begin VB.Label lblMatchCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1755
         End
      End
      Begin VB.PictureBox picGraph 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   0
         ScaleHeight     =   61.648
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   135.731
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   7695
         Begin VB.Label lblCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   26
            Top             =   2640
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label lblCrown 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   0
            Left            =   540
            TabIndex        =   25
            Top             =   2235
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Line linDivide 
            X1              =   0
            X2              =   141.817
            Y1              =   46.567
            Y2              =   46.567
         End
         Begin VB.Label lblBar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FFFFFF&
            Height          =   2415
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   495
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.OptionButton optGraph 
      Caption         =   "Sums"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Graph the sum of all the Traits of one category for the characters: Find the total levels of Influence in the game."
      Top             =   300
      Width           =   1215
   End
   Begin VB.HScrollBar hscScrollGraph 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   140
      Left            =   360
      SmallChange     =   8
      TabIndex        =   18
      Top             =   5640
      Width           =   8415
   End
   Begin VB.CheckBox chkZero 
      Caption         =   "Exclude &zero and (none)"
      Height          =   195
      Left            =   6720
      TabIndex        =   22
      Top             =   1245
      Width           =   2055
   End
   Begin VB.OptionButton optGraph 
      Caption         =   "Maxima"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Graph the highest value of each type of Trait in a given category: Find the highest levels of Influences, Backgrounds, etc."
      Top             =   300
      Width           =   1215
   End
   Begin VB.OptionButton optGraph 
      Caption         =   "Distribution"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Graph the range of one value for the characters: Clan or Tribe populations, distribution of Trait totals, etc."
      Top             =   300
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.ComboBox cboKey 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox cboQuery 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "&Graph the Data"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   255
      Left            =   -360
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.ListBox lstKey 
      Height          =   255
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame fraTraitDistribution 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton optTraitDistribution 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   630
         Width           =   255
      End
      Begin VB.TextBox txtTrait 
         Height          =   315
         Left            =   600
         TabIndex        =   12
         Text            =   "(specific trait)"
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optTraitDistribution 
         Caption         =   "Distinct traits in the category"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optTraitDistribution 
         Caption         =   "Total traits in the &category"
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   45
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.Line linLink 
         Index           =   3
         X1              =   495
         X2              =   240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line linLink 
         Index           =   2
         X1              =   495
         X2              =   240
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line linLink 
         Index           =   1
         X1              =   240
         X2              =   240
         Y1              =   180
         Y2              =   750
      End
      Begin VB.Line linLink 
         Index           =   0
         X1              =   495
         X2              =   0
         Y1              =   750
         Y2              =   750
      End
   End
   Begin VB.Label lblColumnReminder 
      Alignment       =   2  'Center
      Caption         =   "Click columns to see associated characters."
      Height          =   495
      Left            =   6720
      TabIndex        =   35
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Maxima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   1920
      Width           =   5775
   End
   Begin VB.Label lblDistinct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8400
      TabIndex        =   28
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label lblResults 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "          "
      Height          =   195
      Left            =   8325
      TabIndex        =   21
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Graph &Type:"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   315
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "&Data to Graph:"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   1245
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "&Whom to Graph:"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   765
      Width           =   1215
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7680
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OPT_DIST = 0
Private Const OPT_MAXIMA = 1
Private Const OPT_SUM = 2

Private Const OPT_TOTAL = 0
Private Const OPT_DISTINCT = 1
Private Const OPT_SPECIFIC = 2

Private Const ZERO_FAR = 6720
Private Const ZERO_NEAR = 4080

Private Matches As Collection

Private Sub UpdateTitle()
'
' Name:         UpdateTitle
' Description:  Update the title of the graph-to-be.
'

    Dim Title As String

    If optGraph(OPT_DIST).Value Then
        Title = "Distribution of "
        If fraTraitDistribution.Visible Then
            If optTraitDistribution(OPT_TOTAL).Value Then
                Title = Title & "total "
            ElseIf optTraitDistribution(OPT_DISTINCT).Value Then
                Title = Title & "distinct "
            Else
                Title = Title & txtTrait.Text & " "
            End If
        End If
    ElseIf optGraph(OPT_MAXIMA).Value Then
        Title = "Maxima of "
    Else
        Title = "Sums of "
    End If
    
    Title = Title & cboKey.Text & " for " & cboQuery.Text

    lblTitle.Caption = Title

End Sub

Private Sub AddStatisticBar(StatName As String, StatNum As String, MatchList As LinkedTraitList)
'
' Name:         AddStatistic
' Parameters:   StatName        Name of the new statistic bar to add
'               StatNum         Associated statistic number
' Description:  Initialize a new statistic bar.
'
    Dim I As Integer
    Dim X As Integer
    Dim ClipLeft As Boolean
    Dim ClipTop As Boolean
    
    I = lblBar.Count
    
    Load lblBar(I)
    Load lblCaption(I)
    Load lblCrown(I)
    
    lblBar(I).ToolTipText = StatName
    Matches.Add MatchList, "k" & CStr(I)
        
    StatName = " " & StatName
    If picGraph.TextWidth(StatName & " ") >= (3 * lblBar(0).Width) Then
        StatName = StatName & "..."
        Do Until picGraph.TextWidth(StatName & " ") < (3 * lblBar(0).Width) Or StatName = " ..."
            StatName = Left(StatName, Len(StatName) - 4) & "..."
        Loop
    End If
    
    StatName = StatName & " "
    lblCaption(I).Caption = StatName
        
    If I = 1 Then
        lblCaption(I).Top = lblCaption(0).Top
        lblCaption(I).Left = lblBar(0).Left
    Else
        
        lblCaption(I).Left = lblCaption(I - 1).Left + (lblCaption(I - 1).Width / 2) _
                             + lblBar(0).Width - (lblCaption(I).Width / 2)
        
        For X = 1 To I - 1
            ClipLeft = (lblCaption(I).Left < lblCaption(X).Left + lblCaption(X).Width)
            ClipTop = lblCaption(I).Top >= lblCaption(X).Top And _
                      lblCaption(I).Top < (lblCaption(X).Top + lblCaption(X).Height)
            If ClipLeft And ClipTop Then
                lblCaption(I).Top = lblCaption(X).Top + lblCaption(X).Height
                X = 1
            End If
        Next X
        
    End If

    lblBar(I).Caption = StatNum
    lblBar(I).Left = lblCaption(I).Left + (lblCaption(I).Width / 2) - (lblBar(I).Width / 2)
    
    lblCrown(I).Left = lblBar(I).Left
    
    lblCaption(I).BackColor = lblColor(I Mod lblColor.Count).BackColor
    lblBar(I).BackColor = lblColor(I Mod lblColor.Count).BackColor
    
    lblBar(I).ZOrder 1
    lblCaption(I).ZOrder 0

End Sub

Private Sub CompleteStatistics(Stat As StatisticType, Total As Double, High As Double)
'
' Name:         CompleteStatistics
' Parameters:   Stat            The type of statistic performed
'               Total             The 100% value for the entire graph.
'               High            The highest value of all the numbers.
' Description:  Size all the graph columns appropriately, add their numbers, and
'               position the frame within the view port.
'

    Dim Bar As Label
    Dim Increment As Double
    Dim H As Double
    Dim I As Integer
    
    picMatches.Visible = False
    
    If Total > 0 And High > 0 Then
        
        picGraph.Left = 0
        picGraph.Width = lblBar(lblBar.Count - 1).Left + (lblBar(0).Width * 2)
        
        If picGraph.Width > picView.ScaleWidth Then
            hscScrollGraph.Max = Int(picGraph.Width - picView.ScaleWidth) + 1
            hscScrollGraph.Enabled = True
            hscScrollGraph.Value = 0
        Else
            hscScrollGraph.Enabled = False
        End If
        
        Increment = lblBar(0).Height / High
        
        For Each Bar In lblBar
            I = Bar.Index
            If I > 0 Then
                
                Bar.Top = (High - Val(Bar.Caption)) * Increment + lblBar(0).Top
                H = lblCaption(Bar.Index).Top - Bar.Top
                Bar.Height = IIf(H > 0, H, 0)
                Bar.ToolTipText = Bar.ToolTipText & " x" & Bar.Caption _
                                  & ": " & Format(Val(Bar.Caption) / Total, "##0.0%")
                Bar.Caption = "  " & Bar.Caption & "   " & _
                              Format(Val(Bar.Caption) / Total, "##0.0%")
                Bar.Visible = True
                
                lblCrown(I).Top = Bar.Top - lblCrown(I).Height
                If Bar.Top + lblCrown(I).Height > linDivide.Y1 Then
                    lblCrown(I).ToolTipText = Bar.ToolTipText
                    lblCrown(I).Caption = Bar.Caption
                    Bar.Caption = ""
                    lblCrown(I).Visible = True
                End If
                
                lblCaption(I).ToolTipText = Bar.ToolTipText
                lblCaption(I).Visible = True
                
            End If
        Next Bar
        
        Select Case Stat
            Case stDistinctDistribution, stDistribution, stSpecificDistribution
                lblResults.Caption = CStr(Total) & " characters examined"
            Case stMaxima, stSums
                lblResults.Caption = CStr(Total) & " total traits"
        End Select
        
        lblDistinct.Caption = CStr(lblBar.Count - 1) & " distinct values"

        linDivide.X2 = picGraph.Width
        linDivide.ZOrder 0
        
        picGraph.Visible = True

    Else
        
        lblDistinct.Caption = ""
        lblResults.Caption = "No results"
        
    End If

End Sub

Private Sub ClearGraph()
'
' Name:         ClearGraph
' Description:  Unload all the labels, make the graph invisible.
'

    Dim X As Integer

    picMatches.Visible = False
    picView.Cls
    picGraph.Visible = False
    For X = 1 To lblBar.Count - 1
        Unload lblBar(X)
        Unload lblCrown(X)
        Unload lblCaption(X)
    Next X
    hscScrollGraph.Enabled = False
    Do Until Matches.Count = 0
        Matches.Remove 1
    Loop
    
End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initilize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnStatistics
        .SearchName = cboQuery.Text
        .SearchNot = False
        .StatKey = Game.QueryEngine.TitlesToKeys(cboKey.Text)
        .StatTrait = txtTrait.Text
        If optGraph(OPT_DIST).Value Then
            If fraTraitDistribution.Visible Then
                If optTraitDistribution(OPT_SPECIFIC) Then
                    .StatType = stSpecificDistribution
                ElseIf optTraitDistribution(OPT_DISTINCT) Then
                    .StatType = stDistinctDistribution
                Else
                    .StatType = stDistribution
                End If
            Else
                .StatType = stDistribution
            End If
        Else
            If optGraph(OPT_MAXIMA).Value Then
                .StatType = stMaxima
            Else
                .StatType = stSums
            End If
        End If
        .OKZero = (chkZero.Value = vbUnchecked)
        .GameDate = 0
    End With
    
End Sub

Private Sub cboKey_Click()
'
' Name:         cboKey_Click
' Description:  Update the graph title.  Show the trait distribution fields, if needed.
'

    fraTraitDistribution.Visible = (cboKey.ItemData(cboKey.ListIndex) = qtTraitList) And _
        optGraph(OPT_DIST).Value
    chkZero.Left = IIf(fraTraitDistribution.Visible And Not optTraitDistribution(OPT_DISTINCT), _
                       ZERO_FAR, ZERO_NEAR)
    If Not picGraph.Visible Then UpdateTitle
    
End Sub

Private Sub cboQuery_Click()
'
' Name:         cboQuery_Click
' Description:  Update the graph title.
'

    If Not picGraph.Visible Then UpdateTitle
    
End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Unload this form.
'
    Unload Me

End Sub

Private Sub cmdGraph_Click()
'
' Name:         cmdGraph_Click
' Description:  Construct and run the statistics query.  Display the results.
'

    Dim Stat As StatisticType
    Dim Search As QueryClass
    Dim Key As String
    
    Screen.MousePointer = vbHourglass
    
    ClearGraph
    UpdateTitle

    With Game.QueryEngine.QueryList
        .MoveTo cboQuery.Text
        If .Off Then
            Set Search = New QueryClass
            Search.Inventory = qiCharacters
        Else
            Set Search = .Item
        End If
    End With
    
    If optGraph(OPT_DIST).Value Then
        If fraTraitDistribution.Visible Then
            If optTraitDistribution(OPT_SPECIFIC) Then
                Stat = stSpecificDistribution
            ElseIf optTraitDistribution(OPT_DISTINCT) Then
                Stat = stDistinctDistribution
            Else
                Stat = stDistribution
            End If
        Else
            Stat = stDistribution
        End If
    Else
        If optGraph(OPT_MAXIMA).Value Then
            Stat = stMaxima
        Else
            Stat = stSums
        End If
    End If

    Key = Game.QueryEngine.TitlesToKeys(cboKey.Text)

    With Game.QueryEngine
        .GetStatistics Stat, Search, Key, (chkZero.Value = vbUnchecked), txtTrait.Text

        .StatResults.First
        Do Until .StatResults.Off
            AddStatisticBar CStr(.StatResults.Item), CStr(.NumberSet(.StatResults.Item)), _
                            .MatchSet(.StatResults.Item)
            .StatResults.MoveNext
        Loop

        CompleteStatistics Stat, .Total, .Maximum

    End With
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdHide_Click()
'
' Name:         cmdHide_Click
' Description:  Make the list of matching characters vanish.
'
    picMatches.Visible = False
    
End Sub

Private Sub cmdShowCharacter_Click()
'
' Name:         cmdShowCharacter_Click
' Description:  Show the selected character.
'
    
    If lstMatches.ListIndex > -1 Then
    
        Dim MatchNames As LinkedTraitList
        Set MatchNames = Matches(lstMatches.Tag)
        MatchNames.MoveToPlace lstMatches.ListIndex
        If Not MatchNames.Off Then
            mdiMain.ShowCharacterSheet MatchNames.Trait.Name
        End If
        
    End If

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Check to see if the QueryList has changed: if so, repopulate it.
'

    If mdiMain.CheckForChanges(Me, atQueries) Then
    
        Dim Store As String
        
        Store = cboQuery.Text
        cboQuery.Clear
        
        With Game.QueryEngine.QueryList
            .First
            Do Until .Off
                If .Item.Inventory = qiCharacters Then cboQuery.AddItem .Item.Name
                If .Item.Name = Store Then cboQuery.ListIndex = cboQuery.NewIndex
                .MoveNext
            Loop
        End With
    
    End If

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the combo boxes.
'

    Dim Key As Variant

    With Game.QueryEngine.QueryList
        .First
        Do Until .Off
            If .Item.Inventory = qiCharacters Then cboQuery.AddItem .Item.Name
            If .Item.Name = "All Characters" Then cboQuery.ListIndex = cboQuery.NewIndex
            .MoveNext
        Loop
    End With

    With Game.QueryEngine
    
        For Each Key In .TitlesToKeys
        
            If (CInt(.KeysToInventories(Key)) And qiCharacters) Then
                cboKey.AddItem CStr(.KeysToTitles(Key))
                cboKey.ItemData(cboKey.NewIndex) = .KeysToTypes(Key)
                If Key = qkGroup Then cboKey.ListIndex = cboKey.NewIndex
                lstKey.AddItem CStr(.KeysToTitles(Key))
                lstKey.ItemData(lstKey.NewIndex) = .KeysToTypes(Key)
            End If
            
        Next Key
    
    End With

    Set Matches = New Collection

    Me.Show

    picView.Print ""
    picView.Print " Distribution graphs show the number of characters for each value of the"
    picView.Print " aspect you graph.  Use this to find clan or tribe populations, mean Willpower,"
    picView.Print " the ratio of Seelie to Unseelie fae, etc."
    picView.Print ""
    picView.Print " When examining the distribution of a list of traits, there are three uses:"
    picView.Print "    Total Traits: examine the distribution of characters' total Physical "
    picView.Print "       Traits, Influences, numbers of Hekau, etc."
    picView.Print "    Distinct Traits: see a list of all the Disciplines, Flaws, Abilities, "
    picView.Print "       etc. that are presently in your game -- and the numbers of characters"
    picView.Print "       who have them."
    picView.Print "    Specific Traits: How much Brawl or Pure Breed or Underworld is in your"
    picView.Print "       game?  Examining specific traits shows you the number of characters"
    picView.Print "       who have the Trait at x0, x1, x2, and so on."

End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Clear the graph before unloading the form.
'

    ClearGraph
    Set Matches = Nothing

End Sub

Private Sub hscScrollGraph_Change()
'
' Name:         hscScrollGraph_Change
' Description:  Reposition the graph within the view as needed.
'
    picGraph.Left = -hscScrollGraph.Value
    picMatches.Visible = False
    
End Sub

Private Sub lblBar_Click(Index As Integer)
'
' Name:         lblBar_Click
' Description:  Show the associated list of matches.
'
    
    Dim MatchNames As LinkedTraitList
    Dim DisplayName As String
    
    lstMatches.Clear
    lstMatches.Tag = "k" & CStr(Index)
    Set MatchNames = Matches(lstMatches.Tag)
    
    MatchNames.First
    Do Until MatchNames.Off
        DisplayName = MatchNames.Trait.Name
        If MatchNames.Display <> ldSimple Then
            DisplayName = MatchNames.Trait.Total & " / " & DisplayName
        End If
        lstMatches.AddItem DisplayName
        MatchNames.MoveNext
    Loop
    
    If (picGraph.Left + lblBar(Index).Left) < (picView.ScaleWidth / 2) Then
        picMatches.Left = picView.ScaleWidth - picMatches.Width + 0.2
    Else
        picMatches.Left = -0.2
    End If
    lstMatches.BackColor = lblBar(Index).BackColor
    lstMatches.ForeColor = lblBar(Index).ForeColor
    lstMatches.ListIndex = -1
    lblMatchCaption.Caption = lblCaption(Index).Caption
    lblMatchCount.Caption = CStr(lstMatches.ListCount) & " Characters"
    picMatches.Visible = True
    
End Sub

Private Sub lblCaption_Click(Index As Integer)
'
' Name:         lblCaption_Click
' Description:  Show the associated list of matches.
'
    Call lblBar_Click(Index)

End Sub

Private Sub lblCrown_Click(Index As Integer)
'
' Name:         lblCrown_Click
' Description:  Show the associated list of matches.
'
    Call lblBar_Click(Index)
    
End Sub

Private Sub lstMatches_DblClick()
'
' Name:         lstMatches_DblClick
' Description:  Show associated character.
'
    Call cmdShowCharacter_Click
    
End Sub

Private Sub optGraph_Click(Index As Integer)
'
' Name:         optGraph_Click
' Description:  Reformat cboKeys if needed, since maxima and sums only work with traits.
'
    Dim Store As String
    Dim I As Integer
    Dim InfIndex As Integer
        
    Store = cboKey.Text
    
    cboKey.Clear
    For I = 0 To lstKey.ListCount - 1
    
        If (Index = OPT_DIST) Or (lstKey.ItemData(I) = qtTraitList) _
                Or (lstKey.ItemData(I) = qtNumber) Then
            cboKey.AddItem lstKey.List(I)
            cboKey.ItemData(cboKey.NewIndex) = lstKey.ItemData(I)
            If lstKey.List(I) = "Influences" Then InfIndex = cboKey.NewIndex
            If lstKey.List(I) = Store Then cboKey.ListIndex = cboKey.NewIndex
        End If
    
    Next I

    If cboKey.ListIndex = -1 Then cboKey.ListIndex = InfIndex
    chkZero.Left = IIf(fraTraitDistribution.Visible And Not optTraitDistribution(OPT_DISTINCT), _
                       ZERO_FAR, ZERO_NEAR)
    chkZero.Visible = (Index = OPT_DIST) And Not optTraitDistribution(OPT_DISTINCT).Value
    
    If Not picGraph.Visible Then
        picView.Cls
        Select Case Index
            Case OPT_DIST
                picView.Print ""
                picView.Print " Distribution graphs show the number of characters for each value of the"
                picView.Print " aspect you graph.  Use this to find clan or tribe populations, mean Willpower,"
                picView.Print " the ratio of Seelie to Unseelie fae, etc."
                picView.Print ""
                picView.Print " When examining the distribution of a list of traits, there are three uses:"
                picView.Print "    Total Traits: examine the distribution of characters' total Physical "
                picView.Print "       Traits, Influences, numbers of Hekau, etc."
                picView.Print "    Distinct Traits: see a list of all the Disciplines, Flaws, Abilities, "
                picView.Print "       etc. that are presently in your game -- and the numbers of characters"
                picView.Print "       who have them."
                picView.Print "    Specific Traits: How much Brawl or Pure Breed or Underworld is in your"
                picView.Print "       game?  Examining specific traits shows you the number of characters"
                picView.Print "       who have the Trait at x0, x1, x2, and so on."
            Case OPT_MAXIMA
                picView.Print ""
                picView.Print " Graph trait maxima when you're interested in the game's highest existing"
                picView.Print " value for each trait in the selected category.  Use this to keep tabs on the"
                picView.Print " highest levels of Influence, Abilities and the like."
            Case Else 'OPT_SUM
                picView.Print ""
                picView.Print " Sums measure the total combined point values in your game of each distinct"
                picView.Print " trait in the selected category.  Use this to track the the game's total"
                picView.Print " number of traits in each Influence, in Viniculum to each character, etc."
        End Select
    End If
    
End Sub

Private Sub optTraitDistribution_Click(Index As Integer)
'
' Name:         optTraitDistribution_Click
' Description:  Update the title or transfer focus to the text box as needed.
'

    chkZero.Left = ZERO_FAR
    chkZero.Visible = Not (Index = OPT_DISTINCT)

    If Index = OPT_SPECIFIC Then
        txtTrait.SetFocus
    Else
        If Not picGraph.Visible Then UpdateTitle
    End If
    
End Sub

Private Sub picGraph_Click()
'
' Name:         picGraph_Click
' Description:  Hide the list of matches.
'
    picMatches.Visible = False

End Sub

Private Sub picView_Click()
'
' Name:         picGraph_Click
' Description:  Hide the list of matches.
'
    picMatches.Visible = False

End Sub

Private Sub txtTrait_Change()
'
' Name:         txtTrait_Change
' Description:  Select this option, as the user fills it out.
'
    optTraitDistribution(OPT_SPECIFIC).Value = True
    If Not picGraph.Visible Then UpdateTitle

End Sub

Private Sub txtTrait_GotFocus()
'
' Name:         txtTrait_GotFocus
' Description:  Select the control's text.
'
    SelectText txtTrait

End Sub
