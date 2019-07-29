VERSION 5.00
Begin VB.Form frmHealthLevels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Health Level Options"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7305
   Icon            =   "frmHealthLevels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbbreviated 
      Caption         =   "Convert to A&bbreviated health levels"
      Height          =   615
      Left            =   2820
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdExtended 
      Caption         =   "&Convert to E&xtended health levels"
      Height          =   615
      Left            =   660
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.OptionButton optAbbreviated 
      Caption         =   "&Abbreviated health levels (from previous MET releases)"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.OptionButton optExtended 
      Caption         =   "&Extended health levels (from Laws of the Night Revised)"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Value           =   -1  'True
      Width           =   4335
   End
   Begin VB.Label lblLabel 
      Caption         =   "These buttons will convert all characters in the database to the preferred number of health levels."
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   10
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label lblLabel 
      Caption         =   "Conversion"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label lblLabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label lblLabel 
      Caption         =   "This setting affects the number of health levels with which most characters are by default created."
      Height          =   495
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblLabel 
      Caption         =   "Default Health Levels"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label lblLabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmHealthLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ConvertHealth(Extend As Boolean)
'
' Name:         ConvertHealth
' Arguments:    Extend (Boolean) -- TRUE if conversion to extended health levels,
'               FALSE if to abbreviated health levels
' Description:  Convert all the characters in the database to a given base set of
'               health levels
'
    Dim CharForm As Form
    Dim Race As RaceType
    Dim HealthList As LinkedTraitList
    
    '
    ' Unload the visible character sheets
    '
    For Each CharForm In Forms
        If CharForm.Tag = "C" Then Unload CharForm
    Next CharForm
    
    CharacterList.First
    Do Until CharacterList.Off
        
        Race = CharacterList.Item.RaceCode
        
        If Not (Race = gvRaceWraith Or Race = gvRaceVarious) Then
            
            Set HealthList = CharacterList.Item.HealthList
            
            HealthList.MoveTo StdHealth(1)
            
            '
            ' Check the characters by their "Bruised" levels.  Less than 3,
            ' They can be extended; 3 or more, they can be abbreviated.
            '
            '#' If HealthList.Item.Number < 3 And Extend Then
            '#'
            '#'     HealthList.Insert StdHealth(0)
            '#'     HealthList.Insert StdHealth(1)
            '#'     HealthList.Insert StdHealth(1)
            '#'     HealthList.Insert StdHealth(2)
            '#'
            '#' ElseIf HealthList.Item.Number > 2 And Not Extend Then
            '#'
            '#'     HealthList.Remove StdHealth(0)
            '#'     HealthList.Remove StdHealth(1)
            '#'     HealthList.Remove StdHealth(1)
            '#'     HealthList.Remove StdHealth(2)
            '#'
            '#' End If
        
        End If
        
        CharacterList.MoveNext
        
    Loop

    Game.DataChanged = True

End Sub

Private Sub cmdAbbreviated_Click()
'
' Name:         cmdAbbreviated_Click
' Description:  Convert to abbreviated health levels
'
    If MsgBox("This will PERMANENTLY REMOVE the health levels " & vbCrLf & _
            StdHealth(0) & " x1, " & StdHealth(1) & " x2, and " & StdHealth(2) & _
            " x1" & vbCrLf & "from characters that have extended health levels." & vbCrLf & _
            "Wraiths and Various-type characters will not be affected." & vbCrLf & _
            vbCrLf & "Are you sure you want to continue?", vbYesNo, "Convert to Abbreviated Health") _
            = vbYes Then
                        
            Screen.MousePointer = vbHourglass
            ConvertHealth False
            Screen.MousePointer = vbDefault
            
    End If

End Sub

Private Sub cmdExtended_Click()
'
' Name:         cmdExtended_Click
' Description:  Convert to extended health levels
'
    If MsgBox("This will PERMANENTLY ADD the health levels " & vbCrLf & _
            StdHealth(0) & " x1, " & StdHealth(1) & " x2, and " & StdHealth(2) & _
            " x1 to characters without them." & vbCrLf & _
            "Wraiths and Various-type characters will not be affected." & vbCrLf & _
            vbCrLf & "Are you sure you want to continue?", vbYesNo, "Convert to Extended Health") _
            = vbYes Then
            
            Screen.MousePointer = vbHourglass
            ConvertHealth True
            Screen.MousePointer = vbDefault
            
    End If

End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Close window, saving changes.
'
    Game.DataChanged = Game.DataChanged Or (optExtended = Not Game.ExtendedHealth)
    Game.ExtendedHealth = optExtended
    Unload Me
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize option values.
'
    optExtended = Game.ExtendedHealth
    optAbbreviated = Not Game.ExtendedHealth

End Sub
