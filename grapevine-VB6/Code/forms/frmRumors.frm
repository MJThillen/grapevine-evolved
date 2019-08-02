VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmRumors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rumors"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9030
   Icon            =   "frmRumors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   9030
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin TabDlg.SSTab tabTabs 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Date and Options"
      TabPicture(0)   =   "frmRumors.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblRumorCount"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "calCalendar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstDates"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdDelete"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkStatusActive"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkNPC"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmRumors.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstGeneral"
      Tab(1).Control(1)=   "cmdClearAll(0)"
      Tab(1).Control(2)=   "cmdClear(0)"
      Tab(1).Control(3)=   "txtRumor(0)"
      Tab(1).Control(4)=   "lblGeneral"
      Tab(1).Control(5)=   "lblDate(0)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Personal"
      TabPicture(2)   =   "frmRumors.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstCharacters"
      Tab(2).Control(1)=   "cmdClearAll(1)"
      Tab(2).Control(2)=   "cmdClear(1)"
      Tab(2).Control(3)=   "txtRumor(1)"
      Tab(2).Control(4)=   "lblPersonal"
      Tab(2).Control(5)=   "lblDate(1)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Group"
      TabPicture(3)   =   "frmRumors.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstGroups"
      Tab(3).Control(1)=   "cmdClearAll(2)"
      Tab(3).Control(2)=   "cmdClear(2)"
      Tab(3).Control(3)=   "txtRumor(2)"
      Tab(3).Control(4)=   "lblGroup"
      Tab(3).Control(5)=   "lblDate(2)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Influence"
      TabPicture(4)   =   "frmRumors.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdInfluences"
      Tab(4).Control(1)=   "cmdClearAll(3)"
      Tab(4).Control(2)=   "cmdClear(3)"
      Tab(4).Control(3)=   "txtRumor(3)"
      Tab(4).Control(4)=   "lblInfluences"
      Tab(4).Control(5)=   "lblDate(3)"
      Tab(4).ControlCount=   6
      Begin VB.CheckBox chkNPC 
         Caption         =   "Write no rumors for NPCs"
         Height          =   435
         Left            =   480
         TabIndex        =   6
         Top             =   4920
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox chkStatusActive 
         Caption         =   "Only write rumors for characters whose status is ""Active"""
         Height          =   435
         Left            =   480
         TabIndex        =   5
         Top             =   4560
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin MSFlexGridLib.MSFlexGrid grdInfluences 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   27
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   6
         Cols            =   6
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         BorderStyle     =   0
         FormatString    =   "| 1 | 2 | 3 | 4 | 5 ;|Transportation  "
      End
      Begin VB.ListBox lstGeneral 
         Height          =   3825
         IntegralHeight  =   0   'False
         Left            =   -74760
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear all General Rumors"
         Height          =   615
         Index           =   0
         Left            =   -74760
         TabIndex        =   12
         Top             =   4800
         Width           =   2175
      End
      Begin VB.ListBox lstGroups 
         Height          =   3825
         IntegralHeight  =   0   'False
         Left            =   -74760
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   840
         Width           =   2175
      End
      Begin VB.ListBox lstCharacters 
         Height          =   3825
         IntegralHeight  =   0   'False
         Left            =   -74760
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear all Influence Rumors"
         Height          =   615
         Index           =   3
         Left            =   -74760
         TabIndex        =   30
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear all Group Rumors"
         Height          =   615
         Index           =   2
         Left            =   -74760
         TabIndex        =   24
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear all Personal Rumors"
         Height          =   615
         Index           =   1
         Left            =   -74760
         TabIndex        =   18
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear this Rumor"
         Height          =   615
         Index           =   3
         Left            =   -68880
         TabIndex        =   31
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear this Rumor"
         Height          =   615
         Index           =   2
         Left            =   -68880
         TabIndex        =   25
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear this Rumor"
         Height          =   615
         Index           =   1
         Left            =   -68880
         TabIndex        =   19
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear this Rumor"
         Height          =   615
         Index           =   0
         Left            =   -68880
         TabIndex        =   13
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox txtRumor 
         Height          =   3855
         Index           =   3
         Left            =   -71280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox txtRumor 
         Height          =   3855
         Index           =   2
         Left            =   -71280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox txtRumor 
         Height          =   3855
         Index           =   1
         Left            =   -71280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox txtRumor 
         Height          =   3855
         Index           =   0
         Left            =   -71280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   840
         Width           =   4575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Date and its Rumors"
         Height          =   615
         Left            =   5880
         TabIndex        =   7
         Top             =   4800
         Width           =   2415
      End
      Begin VB.ListBox lstDates 
         Height          =   2865
         IntegralHeight  =   0   'False
         Left            =   5880
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin MSACAL.Calendar calCalendar 
         Height          =   3855
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5415
         _Version        =   524288
         _ExtentX        =   9551
         _ExtentY        =   6800
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   1978
         Month           =   1
         Day             =   31
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   -2147483630
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   -2147483630
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   -2147483630
         ValueIsNull     =   -1  'True
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblRumorCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   5880
         TabIndex        =   33
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Date for which to Write Rumors:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   5880
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   " Eligibility "
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   4320
         Width           =   675
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   4440
         Width           =   5295
      End
      Begin VB.Label lblInfluences 
         BackStyle       =   0  'Transparent
         Caption         =   "Influence 1 Rumors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   26
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label lblGroup 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Rumors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblPersonal 
         BackStyle       =   0  'Transparent
         Caption         =   "Rumors for a Character"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblGeneral 
         BackStyle       =   0  'Transparent
         Caption         =   "General Rumors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "January 31, 1998"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -68625
         TabIndex        =   28
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "January 31, 1998"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -68625
         TabIndex        =   22
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "January 31, 1998"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -68625
         TabIndex        =   16
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "January 31, 1998"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -68625
         TabIndex        =   10
         Top             =   480
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmRumors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PresentRumorType
    UpToDate As Boolean
    Category As Integer
    Recipient As String
    Level As Integer
End Type

'tab Constants
Private Const tcDate = 0
Private Const tcGeneral = 1
Private Const tcPersonal = 2
Private Const tcGroup = 3
Private Const tcInfluence = 4

Private Const Checkmark = "×"

Dim RumorList As LinkedRumorList    'Current Rumor List selected from AllRumorLists

Dim PresentRumor As PresentRumorType 'Information about the current rumor and its source
Dim RumorChanging As Boolean        'Disables certain events when change events fire

Dim BuildInProgress As Boolean      'Disables certain events when builds take place
Dim ItemCheckInProgress As Boolean  'Disables the _Click event when an item is checked
Dim UpdateInProgress As Boolean     'Disables some events when calendar updates happen

Private Sub BuildLists(Signal As DataChangedSignalType)
'
' Name:         BuildLists
' Parameters:   Signal      which of the lists to build
' Description:  Uses the CharacterList and RumorList to build the General,
'               Character, Group, and Influence lists.  Clears them then
'               builds from scratch.
'
    
    '
    ' Step 0:  Declare Variables
    '
    Dim GrabValue As String                         'Temporary holder of races and groups
    Dim Races As String                             'String Index of all found races
    Dim Groups As String                            'String Index of all found groups
    
    Dim CharInfluenceList As LinkedTraitList        'A character's .InfluenceList
    Dim MasterInfluenceList As LinkedTraitList      'A master influence list
    Dim CharInfluence As TraitClass                 'A Character's particular influence trait
    Dim MasterInfluence As TraitClass               'An influence trait in the master list
    
    Dim Entry As Integer                            'For iterating trhough the lists
    Dim HasEntries As Boolean                       'General flag for seeing if things are filled
    
    Dim InfluenceFormat As String                   'The format string for the grid's fixed column
    Dim NumberFormat As String                      'The format string for the grid's fixed row
    Dim HighLevel As Integer                         'The level of the highest-level influence around
    Dim GridRow As Integer                          'For iterating through grid rows
    Dim GridCol As Integer                          'For iterating through grid columns
    
    Dim RumorValue As String                        'For storing the Recipient of a rumor
    Dim RumorFound As Boolean                       'Indicated if a rumor was checked
    
    If Not (Signal.CharactersChanged Or Signal.GroupsChanged Or Signal.InfluencesChanged) Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    BuildInProgress = True
    Set MasterInfluenceList = New LinkedTraitList
    
    '
    ' Step 1: Clear the requisite lists and grids
    '
    If Signal.CharactersChanged Then                'Empty the General and Character listboxes
        lstGeneral.Clear
        lstGeneral.AddItem "Public Knowledge"
        lstCharacters.Clear
    End If
    
    If Signal.GroupsChanged Then                    'Empty the Group listbox
        lstGroups.Clear
    End If
    
    If Signal.InfluencesChanged Then                'Empty the Influence grid
        grdInfluences.Cols = 2
        grdInfluences.Rows = 2
        grdInfluences.FormatString = "|;|"
    End If
    
    '
    ' Step 2: Fill the lists and grids
    '
    HighLevel = 0                                                   'nullify the highest influence level
    CharacterList.First
    Do Until CharacterList.Off                                      'For every character in the game--

        If StatusSelected(CharacterList.Item) Then                  '(those that satify the Status
                                                                    ' restrictions, that is)
            If Signal.CharactersChanged Then                        'If we're updating races and names,
                GrabValue = CharacterList.Item.Race
                If InStr(Races, "!~" & GrabValue & "~!") = 0 Then   'Check to see if we have this race;
                    lstGeneral.AddItem GrabValue                    'if not, add it to the listbox.
                    Races = Races & "!~" & GrabValue & "~!"
                End If
                lstCharacters.AddItem CharacterList.Item.Name       'Always add names to the listbox.
            End If
            
            If Signal.GroupsChanged Then                            'if we're updating the groups,
                GrabValue = CharacterList.Item.Group
                If Trim(GrabValue) <> "" And _
                        InStr(Groups, "!~" & GrabValue & "~!") = 0 Then 'Check to see if we have this group;
                    lstGroups.AddItem GrabValue                     'if not, add it to the listbox.
                    Groups = Groups & "!~" & GrabValue & "~!"
                End If
            End If
            
            If Signal.InfluencesChanged Then                        'if we're updating influences,
                                                                    'grab the character's InfluenceList
                Set CharInfluenceList = CharacterList.Item.InfluenceList
                CharInfluenceList.First
                Do Until CharInfluenceList.Off                      'For each influence,
                    Set CharInfluence = CharInfluenceList.Trait
                    MasterInfluenceList.MoveTo CharInfluence.Name
                    If MasterInfluenceList.Off Then                 'Check if it's in the master list.
                        Set MasterInfluence = New TraitClass        'No?  Add a copy.
                        MasterInfluence.Name = CharInfluence.Name
                        '#' MasterInfluence.Number = CharInfluence.Number
                        '#' MasterInfluenceList.InsertTrait MasterInfluence
                                                                    'Yes? If its number is higher, set
                    Else                                            'the master list entry to that number.
                        Set MasterInfluence = MasterInfluenceList.Trait
                        '#' If CharInfluence.Number > MasterInfluence.Number Then _
                                MasterInfluence.Number = CharInfluence.Number
                    End If
                    '#' If CharInfluence.Number > HighLevel Then _
                            HighLevel = CharInfluence.Number
                    CharInfluenceList.MoveNext                      'Next Influence!
                Loop
            
            End If

        End If
        CharacterList.MoveNext                                      'Next Character!
    
    Loop

    '
    ' Step 3:  Set Visibility and build grids
    '
    If Signal.CharactersChanged Then                            'We're working with characters
            
        HasEntries = (Not RumorList Is Nothing)
        lstGeneral.Visible = HasEntries
        txtRumor(rtGeneral).Visible = HasEntries
        cmdClear(rtGeneral).Visible = HasEntries
        If HasEntries Then
            lstGeneral.ListIndex = 0
            lblGeneral = lstGeneral.Text & " Rumors"
        Else
            lblGeneral = "General Rumors"
        End If
        
        HasEntries = (lstCharacters.ListCount > 0 And Not RumorList Is Nothing)
        lstCharacters.Visible = HasEntries
        txtRumor(rtPersonal).Visible = HasEntries
        cmdClear(rtPersonal).Visible = HasEntries
        If HasEntries Then
            lstCharacters.ListIndex = 0
            lblPersonal = "Rumors for " & lstCharacters.Text
        Else
            lblPersonal = "Personal Rumors"
        End If
    
    End If
    
    If Signal.GroupsChanged Then                                'We're Working with Groups

        HasEntries = (lstGroups.ListCount > 0 And Not RumorList Is Nothing)
        lstGroups.Visible = HasEntries
        txtRumor(rtGroup).Visible = HasEntries
        cmdClear(rtGroup).Visible = HasEntries
        If HasEntries Then
            lstGroups.ListIndex = 0
            lblGroup = lstGroups.Text & " Rumors"
        Else
            lblGroup = "Group Rumors"
        End If

    End If
        
    If Signal.InfluencesChanged Then                            'Hoo baby, we're working with Influences!
    
        If MasterInfluenceList.IsEmpty Or RumorList Is Nothing Then 'Are there no influences?
            grdInfluences.Visible = False                       'If no, hide the influence controls
            txtRumor(rtInfluence).Visible = False
            cmdClear(rtInfluence).Visible = False
            lblInfluences = "Influence Rumors"
        Else                                                    'If yes, make them visible, and more...
            grdInfluences.Visible = True
            txtRumor(rtInfluence).Visible = True
            cmdClear(rtInfluence).Visible = True
        
            MasterInfluenceList.First                           'For each master influence,
            Do Until MasterInfluenceList.Off                    'Put its title in a format string
                InfluenceFormat = InfluenceFormat & "|  " _
                        & MasterInfluenceList.Trait.Name
                MasterInfluenceList.MoveNext
            Loop
        
            For Entry = 1 To HighLevel                          'Make a format string for the highest-
                NumberFormat = NumberFormat & "|  " & CStr(Entry)   'level influence
            Next Entry
                                                                'Format the grid with this string
            grdInfluences.FormatString = NumberFormat & ";" & InfluenceFormat
            
            MasterInfluenceList.First                           'Start at the top of the master list;
            lblInfluences.Caption = _
                    MasterInfluenceList.Trait.Name & " 1 Rumors" 'Set the label for the upper-left cell
            
            For GridRow = 1 To grdInfluences.Rows - 1           'For each Row in the grid (and influence
                                                                'in the master list)
                grdInfluences.Row = GridRow                     'Set the row
                For GridCol = 1 To grdInfluences.Cols - 1        'For Each Column in the grid
                    
                    grdInfluences.Col = GridCol                 'Set the Column
                    '#' If GridCol <= MasterInfluenceList.Item.Number Then 'If Influence rumors are needed
                    '#'     With grdInfluences                             'at this level,
                    '#'         .CellBackColor = .BackColor                'Highlight the cell.
                    '#'         .CellForeColor = .ForeColor
                    '#'         .CellFontBold = True
                    '#'     End With
                    '#' Else                                        'Otherwise, if influence rumors aren't
                    '#'     With grdInfluences                      'needed at this level,
                    '#'         .CellBackColor = .BackColorBkg      'un-highlight the cell.
                    '#'         .CellForeColor = .ForeColorFixed
                    '#'     End With
                    '#' End If
                
                Next GridCol                                    'Next Column!
                
                MasterInfluenceList.MoveNext                    'Next Master Influence!
            Next GridRow                                        'Next Row!

            With grdInfluences
                .Col = 1
                .Row = 1
            lblInfluences = .TextMatrix(1, 0) & " 1 Rumors"
                .CellBackColor = .BackColorSel
                .CellForeColor = .ForeColorSel
            End With
            
        End If

    End If
    
    '
    ' Step 4: Check the boxes of those entries with rumors; check the needed squares of the grid.
    '
    
    If Not RumorList Is Nothing Then
    
        RumorList.First
        Do Until RumorList.Off
        
            Entry = 0
            RumorValue = RumorList.Item.Recipient
            RumorFound = False
            Select Case RumorList.Item.Category
                Case rtGeneral
                    Do Until Entry >= lstGeneral.ListCount
                        If lstGeneral.List(Entry) = RumorValue Then
                            lstGeneral.Selected(Entry) = True
                            RumorFound = True
                            Exit Do
                        End If
                        Entry = Entry + 1
                    Loop
                Case rtPersonal
                    Do Until Entry >= lstCharacters.ListCount
                        If lstCharacters.List(Entry) = RumorValue Then
                            lstCharacters.Selected(Entry) = True
                            RumorFound = True
                            Exit Do
                        End If
                        Entry = Entry + 1
                    Loop
                Case rtGroup
                    Do Until Entry >= lstGroups.ListCount
                        If lstGroups.List(Entry) = RumorValue Then
                            lstGroups.Selected(Entry) = True
                            RumorFound = True
                            Exit Do
                        End If
                        Entry = Entry + 1
                    Loop
                Case rtInfluence
                    Entry = 1
                    Do Until Entry >= grdInfluences.Rows
                        If grdInfluences.TextMatrix(Entry, 0) = RumorValue And _
                                RumorList.Item.Level < grdInfluences.Cols Then
                            grdInfluences.Row = Entry
                            grdInfluences.Col = RumorList.Item.Level
                            If grdInfluences.CellBackColor <> grdInfluences.BackColorBkg Then
                                grdInfluences.Text = Checkmark
                                RumorFound = True
                            End If
                        End If
                        Entry = Entry + 1
                    Loop
                    grdInfluences.Col = 1
                    grdInfluences.Row = 1
                    grdInfluences.CellBackColor = grdInfluences.BackColorSel
            End Select
            
'   This code removes all rumors that aren't found; i.e., erases any rumors not
'   presently accessible.  It causes a lot of problems, it seems.
'
'            If Not RumorFound Then
'                RumorList.Remove
'            Else
'                RumorList.MoveNext
'            End If
'
'   Alternate Code:
'
            RumorList.MoveNext

        Loop
    
    End If
    
    '
    ' Cleanup
    '
    Set MasterInfluenceList = Nothing
    BuildInProgress = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub StoreRumor()
'
' Name:         StoreRumor
' Description:  Store the current rumor.
'

    Dim NewText As String
    Dim NewRumor As RumorClass
    Dim List As ListBox
    Dim Entry As Integer
    Dim Cursor As Integer

    If Not PresentRumor.UpToDate And Not RumorList Is Nothing Then
        
        NewText = TrimWhiteSpace(txtRumor(PresentRumor.Category))
        RumorList.MoveTo PresentRumor.Category, _
                PresentRumor.Recipient, PresentRumor.Level
        If NewText <> "" Then
            
            If RumorList.Off Then
                Set NewRumor = New RumorClass
                With NewRumor
                    .Category = PresentRumor.Category
                    .Recipient = PresentRumor.Recipient
                    .Level = PresentRumor.Level
                    .Text = NewText
                End With
                RumorList.InsertSorted NewRumor
            Else
                RumorList.Item.Text = NewText
            End If
            Game.DataChanged = True
            
        Else
            If Not RumorList.Off Then
                RumorList.Remove
                Game.DataChanged = True
            End If
        End If
        
        If PresentRumor.Category <> rtInfluence Then
            
            Select Case PresentRumor.Category
                Case rtGeneral: Set List = lstGeneral
                Case rtPersonal: Set List = lstCharacters
                Case rtGroup: Set List = lstGroups
            End Select
            Entry = 0
            Do Until Entry >= List.ListCount
                If List.List(Entry) = PresentRumor.Recipient Then
                    ItemCheckInProgress = True
                    Cursor = List.ListIndex
                    List.Selected(Entry) = (NewText <> "")
                    List.ListIndex = Cursor
                    ItemCheckInProgress = False
                    Exit Do
                End If
                Entry = Entry + 1
            Loop
        
        Else
        
            Entry = 1
            Do Until Entry >= grdInfluences.Rows
                If grdInfluences.TextMatrix(Entry, 0) = PresentRumor.Recipient Then
                    If NewText <> "" Then
                        grdInfluences.TextMatrix(Entry, PresentRumor.Level) = Checkmark
                    Else
                        grdInfluences.TextMatrix(Entry, PresentRumor.Level) = ""
                    End If
                    Exit Do
                End If
                Entry = Entry + 1
            Loop
        
        End If
        
        PresentRumor.UpToDate = True
    End If

End Sub

Private Sub SetRumor(Category As Integer, Recipient As String, Level As Integer)
'
' Name:         SetRumor
' Parameters:   Category    category of rumor to display
'               Recipient   recipient of rumor to display
'               Level       level of rumor to display
' Description:  Display the rumor for the given category, recipient and level.
'

    PresentRumor.Category = Category
    PresentRumor.Recipient = Recipient
    PresentRumor.Level = Level
    If Not RumorList Is Nothing Then
        RumorList.MoveTo Category, Recipient, Level
        If RumorList.Off Then
            txtRumor(Category) = ""
        Else
            txtRumor(Category) = RumorList.Item.Text
        End If
    End If
    PresentRumor.UpToDate = True

End Sub

Private Function StatusSelected(Character As Object) As Boolean
'
' Name:         StatusSelected
' Parameters:   Character       a character class to check
' Description:  See if this character meets the user's selection criteria
' Returns:      TRUE if it meets the criteria, FALSE otherwise
'

    StatusSelected = True
    If chkStatusActive = vbChecked Then _
        StatusSelected = (LCase(Character.Status) = LCase(ActiveStatus))
    If chkNPC = vbChecked Then _
        StatusSelected = StatusSelected And Not Character.IsNPC
    
End Function

Private Sub calCalendar_AfterUpdate()
'
' Name:         calCalendar_AfterUpdate
' Description:  Acknowledge that some data has changed.
'

    With RumorSignal
        .CharactersChanged = True
        .GroupsChanged = True
        .InfluencesChanged = True
    End With

End Sub

Private Sub calCalendar_BeforeUpdate(Cancel As Integer)
'
' Name:         calCalendar_BeforeUpdate
' Description:  Ensure that the user wants to write rumors for this date.
'               Create a corresponding RumorList.
'

    Dim StringDate As String
    Dim LongDate As Long
    Dim DateIndex As Integer
    Dim NewList As LinkedRumorList
    Dim LabelDate As Label
    
    UpdateInProgress = True
    LongDate = CLng(calCalendar.Value)
    
    AllRumorLists.MoveTo CStr(calCalendar.Value)
    If AllRumorLists.Off Then
    
        StringDate = Format(calCalendar.Value, "mmmm d, yyyy")
        If MsgBox("Do you want to write rumors for " & StringDate & "?", _
                vbYesNo Or vbDefaultButton2, "Add Date") = vbYes Then
                
            DateIndex = 0
            Do Until DateIndex >= lstDates.ListCount
                If LongDate < lstDates.ItemData(DateIndex) Then Exit Do
                DateIndex = DateIndex + 1
            Loop
            lstDates.AddItem StringDate, DateIndex
            lstDates.ItemData(DateIndex) = LongDate
            lstDates.ListIndex = DateIndex
            
            For Each LabelDate In lblDate
                LabelDate = StringDate
            Next LabelDate
            lblRumorCount = "0 rumor(s) have been written for " & lstDates.Text
            
            Set NewList = New LinkedRumorList
            NewList.DateStamp = calCalendar.Value
            AllRumorLists.Append NewList
            Set RumorList = NewList
            
            Set NewList = New LinkedRumorList
            NewList.DateStamp = calCalendar.Value
            InfluenceUseList.Append NewList
            DatesChangedInfUse = True
            
            Game.DataChanged = True
            
            With RumorSignal
                .CharactersChanged = True
                .GroupsChanged = True
                .InfluencesChanged = True
            End With
            
        Else
            Cancel = True
        End If
        
    Else

        DateIndex = -1
        Do Until DateIndex = lstDates.ListCount - 1
            DateIndex = DateIndex + 1
            If LongDate = lstDates.ItemData(DateIndex) Then Exit Do
        Loop
        lstDates.ListIndex = DateIndex
        Set RumorList = AllRumorLists.Item
        lblRumorCount = CStr(RumorList.Count) & " rumor(s) have been written for " & lstDates.Text

    End If
    UpdateInProgress = False
    
End Sub

Private Sub calCalendar_Click()
'
' Name:         calCalendar_Click
' Description:  Call calCalendar_BeforeUpdate.
'

    Dim Cancel As Integer

    If lstDates.ListCount = 0 Then
        Call calCalendar_BeforeUpdate(Cancel)
    End If

End Sub

Private Sub chkNPC_Click()
'
' Name:         chkNPC_Click
' Description:  Acknowledge that data has changed; store this setting in the registry.
'

    With RumorSignal
        .CharactersChanged = True
        .GroupsChanged = True
        .InfluencesChanged = True
    End With

    SaveSetting App.Title, "RumorOptions", "NPCs", chkNPC.Value

End Sub

Private Sub chkStatusActive_Click()
'
' Name:         chkStatusActive_Click
' Description:  Acknowledge that data has changed; store this setting in the registry.
'
    
    With RumorSignal
        .CharactersChanged = True
        .GroupsChanged = True
        .InfluencesChanged = True
    End With

    SaveSetting App.Title, "RumorOptions", "Active", chkStatusActive.Value

End Sub

Private Sub cmdClear_Click(Index As Integer)
'
' Name:         cmdClear_Click
' Description:  Clear the current rumor.
'

    If MsgBox("Are you sure you want to clear this rumor?", vbYesNo, "Clear Rumor") = vbYes Then
        txtRumor(Index) = ""
        PresentRumor.UpToDate = False
        StoreRumor
    End If
    
End Sub

Private Sub cmdClearAll_Click(Index As Integer)
'
' Name:         cmdClearAll_Click
' Description:  Clear all the rumors from this category.
'

    Dim X As Integer
    Dim Y As Integer
    Dim List As ListBox

    If MsgBox("Are you sure you want to clear all rumors for this category?" & vbCrLf & _
            "This will PERMANENTLY remove them from the database.", vbYesNo, _
            "Clear All") = vbYes Then
    
        If Not RumorList Is Nothing Then
        
            txtRumor(Index) = ""
            PresentRumor.UpToDate = False
            StoreRumor
            
            RumorList.First
            Do Until RumorList.Off
                If RumorList.Item.Category = Index Then
                    RumorList.Remove
                Else
                    RumorList.MoveNext
                End If
            Loop
            Game.DataChanged = True
            
            If Index <> rtInfluence Then
                ItemCheckInProgress = True
                Select Case Index
                    Case rtGeneral: Set List = lstGeneral
                    Case rtPersonal: Set List = lstCharacters
                    Case rtGroup: Set List = lstGroups
                End Select
                Y = List.ListIndex
                For X = 0 To List.ListCount - 1
                    List.Selected(X) = False
                Next X
                List.ListIndex = Y
                ItemCheckInProgress = False
            Else
                For X = 1 To grdInfluences.Rows - 1
                    For Y = 1 To grdInfluences.Cols - 1
                        grdInfluences.TextMatrix(X, Y) = ""
                    Next Y
                Next X
            End If
    
        End If

    End If

End Sub

Private Sub cmdDelete_Click()
'
' Name:         cmdDelete_Click
' Description:  Delete the selected rumor date.
'

    Dim Answer As Integer
    Dim Cursor As Integer
    
    If lstDates.ListIndex > -1 Then
    
        Answer = vbYes
        InfluenceUseList.MoveTo RumorList.Name
        
        If Not RumorList Is Nothing Then
            If RumorList.Count > 0 Or InfluenceUseList.Item.Count > 0 Then _
                    Answer = MsgBox("Are you sure you want to delete " _
                    & lstDates.Text & "?" & vbCrLf & "This will PERMANENTLY remove " _
                    & "all rumors and records of influence use associated with this date!", _
                    vbDefaultButton2 Or vbYesNo, "Delete Date")
        End If
        
        If Answer = vbYes Then
            InfluenceUseList.Remove
            DatesChangedInfUse = True
            AllRumorLists.MoveTo RumorList.Name
            AllRumorLists.Remove
            Game.DataChanged = True
            Cursor = lstDates.ListIndex
            lstDates.RemoveItem Cursor
            If Cursor >= lstDates.ListCount Then Cursor = lstDates.ListCount - 1
            lstDates.ListIndex = Cursor
        End If
        
    End If

End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'

    Unload Me

End Sub

Private Sub Form_Activate()
'
' Name:         Form_Activate
' Description:  Rebuild any lists that need it.
'

    If tabTabs.Tab <> tcDate Then
        Call BuildLists(RumorSignal)
        With RumorSignal
            .CharactersChanged = False
            .GroupsChanged = False
            .InfluencesChanged = False
        End With
    End If

End Sub

Private Sub Form_Deactivate()
'
' Name:         Form_Deactivate
' Description:  Call StoreRumor.
'

    StoreRumor

End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the controls and data.
'

    Dim StringDate As String
    Dim LongDate As Long
    Dim Entry As Integer

    If Not AllRumorLists.IsEmpty Then
            
        AllRumorLists.First
        lstDates.AddItem Format(AllRumorLists.Item.DateStamp, "mmmm d, yyyy")
        lstDates.ItemData(0) = CLng(AllRumorLists.Item.DateStamp)
        AllRumorLists.MoveNext
        
        Do Until AllRumorLists.Off
            LongDate = CLng(AllRumorLists.Item.DateStamp)
            Entry = 0
            Do Until Entry >= lstDates.ListCount
                If lstDates.ItemData(Entry) < LongDate Then
                    Entry = Entry + 1
                Else
                    Exit Do
                End If
            Loop
            lstDates.AddItem Format(AllRumorLists.Item.DateStamp, "mmmm d, yyyy"), Entry
            lstDates.ItemData(Entry) = LongDate
            AllRumorLists.MoveNext
        Loop
        
    End If
        
    calCalendar.Month = Month(Date)
    calCalendar.Year = Year(Date)
    calCalendar.Value = Empty
    
    If lstDates.ListCount > 0 Then lstDates.ListIndex = lstDates.ListCount - 1
    
    PresentRumor.UpToDate = True
    
    chkStatusActive = GetSetting(App.Title, "RumorOptions", "Active", vbChecked)
    chkNPC = GetSetting(App.Title, "RumorOptions", "NPCs", vbChecked)
    
    With RumorSignal
        .CharactersChanged = True
        .GroupsChanged = True
        .InfluencesChanged = True
    End With
    Call BuildLists(RumorSignal)
    With RumorSignal
        .CharactersChanged = False
        .GroupsChanged = False
        .InfluencesChanged = False
    End With
    mdiMain.OrientForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Call StoreRumor.
'

    StoreRumor

End Sub

Private Sub grdInfluences_EnterCell()
'
' Name:         grdInfluences_EnterCell
' Description:  Set the diplayed rumor to the influence and level just selected.
'
    
    If Not BuildInProgress Then
        With grdInfluences
            If .CellBackColor = .BackColorBkg And .Col > 1 Then
                .Col = .Col - 1
            Else
    
                StoreRumor
                lblInfluences = .TextMatrix(.Row, 0) & " " & CStr(.Col) & " Rumors"
                .CellBackColor = .BackColorSel
                .CellForeColor = .ForeColorSel
                SetRumor rtInfluence, _
                        grdInfluences.TextMatrix(grdInfluences.Row, 0), _
                        grdInfluences.Col

            End If
        End With
    End If
    
End Sub

Private Sub grdInfluences_KeyPress(KeyAscii As Integer)
'
' Name:         grdInfluences_KeyPress
' Description:  Begin typing in the rumor field.
'
    
    If txtRumor(3).Visible Then
        txtRumor(3).SetFocus
        txtRumor(3).SelStart = Len(txtRumor(3))
        SendKeys Chr(KeyAscii)
    End If
    
End Sub

Private Sub grdInfluences_LeaveCell()
'
' Name:         grdInfluences_LeaveCell
' Description:  Deselect this grid cell.
'

    With grdInfluences
        If Not BuildInProgress And .CellBackColor = .BackColorSel Then
            .CellBackColor = .BackColor
            .CellForeColor = .ForeColor
        End If
    End With

End Sub

Private Sub lstCharacters_Click()
'
' Name:         lstCharacters_Click
' Description:  Display the personal rumors for this character.
'

    If Not (ItemCheckInProgress Or BuildInProgress) Then
        StoreRumor
        lblPersonal = "Rumors for " & lstCharacters.Text
        SetRumor rtPersonal, lstCharacters.Text, 0
    End If
    
End Sub

Private Sub lstCharacters_ItemCheck(Item As Integer)
'
' Name:         lstCharacters_ItemCheck
' Description:  Disable checking items.
'

    If Not (BuildInProgress Or ItemCheckInProgress) Then
        ItemCheckInProgress = True
        lstCharacters.Selected(Item) = Not lstCharacters.Selected(Item)
        ItemCheckInProgress = False
    End If
    
End Sub

Private Sub lstCharacters_KeyPress(KeyAscii As Integer)
'
' Name:         lstCharacters_KeyPress
' Description:  Begin typing in the rumor text field.
'
    
    If txtRumor(1).Visible Then
        txtRumor(1).SetFocus
        txtRumor(1).SelStart = Len(txtRumor(1))
        SendKeys Chr(KeyAscii)
    End If
    
End Sub

Private Sub lstDates_Click()
'
' Name:         lstDates_Click
' Description:  Select a rumor date.  Report the number of rumors written for
'               that date.
'

    Dim LabelDate As Label
    Dim SameDate As Boolean

    If Not UpdateInProgress Then
    
        If Not RumorList Is Nothing Then
            SameDate = (RumorList.DateStamp = CDate(lstDates.Text))
        End If
        
        If Not SameDate Then
            AllRumorLists.MoveTo CDate(lstDates.Text)
            If AllRumorLists.Off Then
                MsgBox "RumorList matching error!"
                Set RumorList = Empty
            Else
                Set RumorList = AllRumorLists.Item
                With RumorSignal
                    .CharactersChanged = True
                    .GroupsChanged = True
                    .InfluencesChanged = True
                End With
            End If
        End If
        
        calCalendar.Value = RumorList.DateStamp
        For Each LabelDate In lblDate
            LabelDate = Format(calCalendar.Value, "mmmm d, yyyy")
        Next LabelDate
        lblRumorCount = CStr(RumorList.Count) & " rumor(s) have been written for " & lstDates.Text
        
    End If

End Sub

Private Sub lstGeneral_Click()
'
' Name:         lstGeneral_Click
' Description:  Display the selected general rumor.
'
    
    If Not (ItemCheckInProgress Or BuildInProgress) Then
        StoreRumor
        lblGeneral = lstGeneral.Text & " Rumors"
        SetRumor rtGeneral, lstGeneral.Text, 0
    End If
    
End Sub

Private Sub lstGeneral_ItemCheck(Item As Integer)
'
' Name:         lstGeneral_ItemCheck
' Description:  Disable checking items.
'

    If Not (BuildInProgress Or ItemCheckInProgress) Then
        ItemCheckInProgress = True
        lstGeneral.Selected(Item) = Not lstGeneral.Selected(Item)
        ItemCheckInProgress = False
    End If
    
End Sub

Private Sub lstGeneral_KeyPress(KeyAscii As Integer)
'
' Name:         lstGeneral_KeyPress
' Description:  Begin typing in the rumor text field.
'

    If txtRumor(0).Visible Then
        txtRumor(0).SetFocus
        txtRumor(0).SelStart = Len(txtRumor(0))
        SendKeys Chr(KeyAscii)
    End If
    
End Sub

Private Sub lstGroups_Click()
'
' Name:         lstGroups_Click
' Description:  Display the rumor for the selected group.
'
    
    If Not (ItemCheckInProgress Or BuildInProgress) Then
        StoreRumor
        lblGroup = lstGroups.Text & " Rumors"
        SetRumor rtGroup, lstGroups.Text, 0
    End If
    
End Sub

Private Sub lstGroups_ItemCheck(Item As Integer)
'
' Name:         lstGroups_ItemCheck
' Description:  Disable checking items.
'

    If Not (BuildInProgress Or ItemCheckInProgress) Then
        ItemCheckInProgress = True
        lstGroups.Selected(Item) = Not lstGroups.Selected(Item)
        ItemCheckInProgress = False
    End If
    
End Sub

Private Sub lstGroups_KeyPress(KeyAscii As Integer)
'
' Name:         lstGroups_KeyPress
' Description:  Begin typing in the rumor text field.
'

    If txtRumor(2).Visible Then
        txtRumor(2).SetFocus
        txtRumor(2).SelStart = Len(txtRumor(2))
        SendKeys Chr(KeyAscii)
    End If
    
End Sub

Private Sub tabTabs_Click(PreviousTab As Integer)
'
' Name:         tabTabs_Click
' Description:  Display the appropriate rumor for the selected tab.
'

    If PreviousTab <> tabTabs.Tab Then
        Call StoreRumor
        Call Form_Activate
        If tabTabs.Tab <> tcDate And Not RumorList Is Nothing Then
            Select Case tabTabs.Tab     'This will end up harmless if things aren't visible; the
                Case tcGeneral          'change events will never fire and set the rumors out-of-date
                    SetRumor rtGeneral, lstGeneral.Text, 0
                Case tcPersonal
                    SetRumor rtPersonal, lstCharacters.Text, 0
                Case tcGroup
                    SetRumor rtGroup, lstGroups.Text, 0
                Case tcInfluence
                    SetRumor rtInfluence, _
                            grdInfluences.TextMatrix(grdInfluences.Row, 0), _
                            grdInfluences.Col
            End Select
            
        ElseIf Not RumorList Is Nothing Then
            lblRumorCount = CStr(RumorList.Count) & " rumor(s) have been written for " & lstDates.Text
        End If
        
    End If
    
End Sub

Private Sub txtRumor_Change(Index As Integer)
'
' Name:         txtRumor_Change
' Description:  Remember that this rumor has not yet been recorded.
'

    PresentRumor.UpToDate = False
    
End Sub
