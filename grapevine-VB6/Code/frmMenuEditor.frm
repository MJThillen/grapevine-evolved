VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenuEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grapevine Menu Editor"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmMenuEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9030
   Begin VB.Frame fraMenuFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4035
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   2715
      Begin MSComctlLib.ListView lvwMenu 
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
         EndProperty
      End
      Begin VB.ComboBox cboView 
         Height          =   315
         ItemData        =   "frmMenuEditor.frx":058A
         Left            =   600
         List            =   "frmMenuEditor.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3660
         Width           =   2055
      End
      Begin VB.Label lblView 
         Caption         =   "&View:"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblMenuCaption 
         Caption         =   "Grapevine &Menus:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Frame fraFileFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   8295
      Begin VB.CommandButton cmdMergeMenu 
         Caption         =   "Merge/Update Menu..."
         Height          =   375
         Left            =   2880
         TabIndex        =   66
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtDescription 
         Height          =   2175
         Left            =   4080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdLoadMenu 
         Caption         =   "&Load Menu..."
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmdSaveMenu 
         Caption         =   "&Save Menu"
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CommandButton cmdSaveMenuAs 
         Caption         =   "Save Menu &As..."
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "File &Description"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   11
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Full Path"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   9
         Top             =   405
         Width           =   1335
      End
   End
   Begin VB.Frame fraItemFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   8295
      Begin VB.Frame fraMenu 
         Caption         =   "Menu Attributes"
         Height          =   3255
         Left            =   2880
         TabIndex        =   39
         Top             =   120
         Width           =   5295
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmMenuEditor.frx":058E
            Left            =   1440
            List            =   "frmMenuEditor.frx":0590
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chkRequired 
            Caption         =   "&Required"
            Height          =   255
            Left            =   3360
            TabIndex        =   47
            Top             =   1800
            Width           =   1695
         End
         Begin VB.ComboBox cboDisplay 
            Height          =   315
            ItemData        =   "frmMenuEditor.frx":0592
            Left            =   1440
            List            =   "frmMenuEditor.frx":0594
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   2160
            Width           =   2415
         End
         Begin VB.CheckBox chkAddNote 
            Caption         =   "Add Note &with Item"
            Height          =   255
            Left            =   1440
            TabIndex        =   46
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CheckBox chkNegative 
            Caption         =   "Ne&gative"
            Height          =   435
            Left            =   3360
            TabIndex        =   45
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CheckBox chkAlphabetized 
            Caption         =   "Alphabeti&zed"
            Height          =   435
            Left            =   1440
            TabIndex        =   44
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtMenuName 
            Height          =   375
            Left            =   1440
            TabIndex        =   41
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Double-click the menu name at left to open and edit its contents."
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   50
            Top             =   2760
            Width           =   4935
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "&Category"
            Height          =   285
            Index           =   13
            Left            =   360
            TabIndex        =   42
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Di&splay"
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   48
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "N&ame"
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   40
            Top             =   525
            Width           =   975
         End
      End
      Begin VB.Frame fraIncludeMenu 
         Caption         =   "Menu Inclusion"
         Height          =   2775
         Left            =   2880
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   5295
         Begin VB.CommandButton cmdGoToIncludeMenu 
            Caption         =   "&Go To"
            Height          =   375
            Left            =   3960
            TabIndex        =   24
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtIncludeName 
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   480
            Width           =   3615
         End
         Begin VB.ListBox lstIncludeMenu 
            Height          =   1425
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "N&ame"
            Height          =   285
            Index           =   10
            Left            =   360
            TabIndex        =   20
            Top             =   525
            Width           =   975
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "In&cluded Menu"
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   22
            Top             =   990
            Width           =   1095
         End
      End
      Begin VB.Frame fraItem 
         Caption         =   "Menu Item Attributes"
         Height          =   2775
         Left            =   2880
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   5295
         Begin MSComCtl2.UpDown updCost 
            Height          =   375
            Left            =   2055
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            Max             =   1
            Min             =   -1
            Orientation     =   1
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtItemNote 
            Height          =   375
            Left            =   1440
            TabIndex        =   31
            Top             =   1440
            Width           =   3615
         End
         Begin VB.TextBox txtItemCost 
            Height          =   375
            Left            =   1440
            TabIndex        =   29
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtItemName 
            Height          =   375
            Left            =   1440
            TabIndex        =   27
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "N&ote"
            Height          =   285
            Index           =   9
            Left            =   360
            TabIndex        =   30
            Top             =   1485
            Width           =   975
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "&Cost"
            Height          =   285
            Index           =   8
            Left            =   360
            TabIndex        =   28
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "N&ame"
            Height          =   285
            Index           =   7
            Left            =   360
            TabIndex        =   26
            Top             =   525
            Width           =   975
         End
      End
      Begin VB.Frame fraSubmenu 
         Caption         =   "Submenu Link Attributes"
         Height          =   2775
         Left            =   2880
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   5295
         Begin VB.CommandButton cmdGoToSubmenu 
            Caption         =   "&Go To"
            Height          =   375
            Left            =   3960
            TabIndex        =   38
            Top             =   960
            Width           =   1095
         End
         Begin VB.ListBox lstLinkedMenu 
            Height          =   1425
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtSubmenuName 
            Height          =   375
            Left            =   1440
            TabIndex        =   35
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Lin&ked Menu"
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   36
            Top             =   990
            Width           =   975
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "N&ame"
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   34
            Top             =   525
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdNewMenu 
         Caption         =   "&New Menu"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         ToolTipText     =   "Create a new menu."
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CommandButton cmdIncludeMenu 
         Caption         =   "Incl&ude Menu"
         Height          =   375
         Left            =   5640
         TabIndex        =   18
         ToolTipText     =   "Include the contents of another menu in the current menu."
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmdNewSubmenu 
         Caption         =   "New Submenu &Link"
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         ToolTipText     =   "Create a new submenu in the current menu.  A submenu links to another menu."
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmdNewMenuItem 
         Caption         =   "New Menu It&em"
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         ToolTipText     =   "Create a new menu item in the current menu, before the selected item. "
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Tag             =   "Del"
         ToolTipText     =   "Delete the currently selected item."
         Top             =   4080
         Width           =   2535
      End
      Begin MSComCtl2.UpDown updPosition 
         Height          =   375
         Left            =   2880
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label lblPosition 
         Caption         =   "&Position"
         Height          =   255
         Left            =   3120
         TabIndex        =   63
         Top             =   3075
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame fraToolFrame 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   360
      TabIndex        =   51
      Top             =   1200
      Width           =   8295
      Begin VB.TextBox txtShortcuts 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   1815
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   52
         Text            =   "frmMenuEditor.frx":0596
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmdCopyMenu 
         Caption         =   "Cop&y Menu"
         Height          =   375
         Left            =   5520
         TabIndex        =   59
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton cmdFivePower 
         Caption         =   "Create Five-&Power Menu"
         Height          =   375
         Left            =   5520
         TabIndex        =   58
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   5520
         TabIndex        =   55
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   4080
         TabIndex        =   54
         Top             =   360
         Width           =   3975
      End
      Begin VB.CheckBox chkConfirmDelete 
         Caption         =   "&Confirm Deletions"
         Height          =   255
         Left            =   2880
         TabIndex        =   60
         Tag             =   "Del"
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.OptionButton optExact 
         Caption         =   "&Exact Match"
         Height          =   375
         Left            =   4080
         TabIndex        =   57
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton optContains 
         Caption         =   "Contains Te&xt"
         Height          =   375
         Left            =   4080
         TabIndex        =   56
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Find &What:"
         Height          =   255
         Index           =   12
         Left            =   2880
         TabIndex        =   53
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdESCClose 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   -360
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   8280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".gvm"
      Filter          =   "Grapevine Menu Files (*.gvm)|*.gvm;*.gcm|All Files (*.*)|*.*"
      MaxFileSize     =   2048
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   5055
      Left            =   240
      TabIndex        =   62
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&File Information"
            Key             =   "File"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Menu &Items"
            Key             =   "Items"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Menu &Tools"
            Key             =   "Tools"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmMenuEditor.frx":0686
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblFilename 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   61
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurrentTabKey As String         'the current tab
Private MenuSet As MenuSetClass         'list of all menus
Private CurrentMenu As LinkedMenuList   'currently selected menu
Private CurrentItem As ListItem         'currently selected list item
Private InsideMenu As Boolean           'whether we're browsing inside or outside a menu
Private BackSelection As Boolean
Private Populating As Boolean           'whether the program is currently populating data

Private Sub CreateNewMenuEntry(BaseName As String, Cost As String, Note As String)
'
' Name:         CreateNewMenuEntry
' Parameters:   BaseName            base name for the new menu entry
'               Cost                its cost
'               Note                its note
' Description:  Create a new menu item, submenu link, or include menu in the
'               current menu.
'

    Dim NewItemName As String
    Dim NewItemNode As Node
    
    If Not CurrentMenu Is Nothing Then
        NewItemName = CreateNewName(CurrentMenu, BaseName)
        If InsideMenu And Not CurrentItem Is Nothing Then
            If CurrentItem.Tag = "(back)" Then
                CurrentMenu.First
            Else
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.MoveNext
            End If
        Else
            CurrentMenu.Last
            CurrentMenu.MoveNext
        End If
        CurrentMenu.Insert NewItemName, Cost, Note
        RefreshMenuItems CurrentMenu, NewItemName
        MenuSet.DataChanged = True
    End If
    
End Sub
Private Function CreateNewName(List As Object, Base As String) As String
'
' Name:         CreateNewItemName
' Parameters:   List            the list to check for the new name
'               Base            the root of the new name
' Description:  Return a unique name for inclusion in the given list.
' Return:       the new name
'

    Dim I As Integer
    
    I = 0
    CreateNewName = Base
    List.MoveTo CreateNewName
    Do Until List.Off()
        I = I + 1
        CreateNewName = Base & CStr(I)
        List.MoveTo CreateNewName
    Loop

End Function

Private Sub PromptForSave(ByRef Continue As Boolean)
'
' Name:         PromptForSave
' Parameters:   Continue        set to whether user wants to cancel
' Description:  If the menu data has changed, ask if the user wants to save.  If so,
'               call cmdSaveMenu_Click.
'

    Dim Answer As Integer
        
    ValidateControls
    
    If MenuSet.DataChanged Then
        Answer = MsgBox("Do you want to save the menu file first?", _
                 vbYesNoCancel + vbQuestion, "Save Menus?")
        Select Case Answer
            Case vbYes
                Call cmdSaveMenu_Click
                Continue = True
            Case vbNo
                Continue = True
            Case vbCancel
                Continue = False
        End Select
    Else
        Continue = True
    End If

End Sub

Private Function ValidateName(Box As TextBox, List As Object, Original As String) As Boolean
'
' Name:         ValidateName
' Parameters:   Box             text box to validate
'               List            LinkedMenuList or MenuSet object to find duplicates with
'               Original        the previous value
'               WaitForName     whether to hold the user until
'                               he supplies valid input
' Description:  returns whether a string contains no patterns that
'               will confuse Grapevine.
' Returns:      Whether Item is a good menu or menu item name.
'

    ValidateName = Not (Box.Text = "(back)" Or Box.Text Like "*:" Or Box.Text Like "*+" Or _
                        Box.Text Like "* (*)" Or Box.Text Like "* x*#")
    
    If Not ValidateName Then
        MsgBox "The text you've entered contains special characters that" & vbCrLf & _
               "will confuse Grapevine.  Do not use parentheses, multipliers," & vbCrLf & _
               "or a terminating colon or plus sign in your text.", vbOKOnly + vbExclamation, _
               "Bad Text"
    Else
    
        List.MoveTo Box.Text
        ValidateName = List.Off() And LCase(Original) <> LCase(Box.Text)
        If Not ValidateName Then
            MsgBox "The name you entered is already being used.  Please enter another.", _
                   vbOKOnly + vbExclamation, "Duplicate Name"
        End If
    
    End If
    
    If Not ValidateName Then
        Box.Text = Original
        SelectText Box
    End If
    
End Function

Public Sub RefreshMenus(Optional SelMenu As LinkedMenuList = Nothing)
'
' Name:         RefreshMenus
' Parameters:   SelMenu      Menu to select at the end.
'                               If nothing, don't change the selection
' Description:  Repopulate the list of menu items.
'

    Dim aMenu As LinkedMenuList
    Dim NewItem As ListItem
    Dim NodeText As String
    Dim ShowCategory As Long
    Dim Found As Boolean
    Dim NextSelIndex As Long
    Dim SelName As String
    Dim ViewCat As Long
    
    If SelMenu Is Nothing Then
        If Not lvwMenu.SelectedItem Is Nothing Then SelName = lvwMenu.SelectedItem.Key
        ShowCategory = cboView.ItemData(cboView.ListIndex)
    Else
        SelName = "k" & SelMenu.Name
        ShowCategory = SelMenu.Category
        If cboView.ItemData(cboView.ListIndex) = gvRaceNone Then ShowCategory = gvRaceNone
    End If
    NextSelIndex = 1
    If Not lvwMenu.SelectedItem Is Nothing Then NextSelIndex = lvwMenu.SelectedItem.Index + 1
    
    ViewCat = cboView.ItemData(cboView.ListIndex)
    If Not (ViewCat = gvRaceNone Or ShowCategory = ViewCat) Then
        Populating = True
        cboView.ListIndex = 0
        For ViewCat = 1 To cboView.ListCount - 1
            If cboView.ItemData(ViewCat) = ShowCategory Then cboView.ListIndex = ViewCat
        Next ViewCat
        Populating = False
    End If
    
    Set CurrentItem = Nothing
    lvwMenu.ListItems.Clear
    MenuSet.First
    InsideMenu = False
    lblMenuCaption.Caption = "Grapevine &Menus:"
    cboView.Visible = True
    lblView.Visible = True
    
    Do Until MenuSet.Off
        Set aMenu = MenuSet.Menu
        If (aMenu.Category = ShowCategory) Or (ShowCategory = gvRaceNone) Then
            Set NewItem = lvwMenu.ListItems.Add(Text:=aMenu.Name & ":", Key:="k" & aMenu.Name)
            NewItem.Tag = aMenu.Name
            If NewItem.Key = SelName Then Found = True
        End If
        MenuSet.MoveNext
    Loop
    
    If Found Then
        Set lvwMenu.SelectedItem = lvwMenu.ListItems(SelName)
    ElseIf lvwMenu.ListItems.Count >= NextSelIndex Then
        Set lvwMenu.SelectedItem = lvwMenu.ListItems(NextSelIndex)
    ElseIf lvwMenu.ListItems.Count > 0 Then
        Set lvwMenu.SelectedItem = lvwMenu.ListItems(1)
    End If
    If Not lvwMenu.SelectedItem Is Nothing Then lvwMenu.SelectedItem.EnsureVisible
    Call lvwMenu_ItemClick(lvwMenu.SelectedItem)
    
End Sub

Public Sub RefreshMenuItems(SelMenu As LinkedMenuList, Optional SelName As String = "")
'
' Name:         RefreshMenuItems
' Parameters:   SelMenu         menu to display
'               SelName         item to select
' Description:  Repopulate the list of menu items.
'

    Dim aMenu As LinkedMenuList
    Dim NewItem As ListItem
    Dim NodeText As String
    Dim ShowCategory As RaceType
    Dim Found As Boolean
    Dim NextSelIndex As Long
    Dim CurrCat As Long
    
    NextSelIndex = 1
    If Not lvwMenu.SelectedItem Is Nothing Then NextSelIndex = lvwMenu.SelectedItem.Index + 1
    If SelName = "" Then
        If Not lvwMenu.SelectedItem Is Nothing Then SelName = lvwMenu.SelectedItem.Key
    ElseIf SelName = "1" Then
        SelName = ""
        NextSelIndex = 2
    Else
        SelName = "k" & SelName
    End If
        
    Set CurrentMenu = SelMenu
    CurrCat = cboView.ItemData(cboView.ListIndex)
    If Not (CurrCat = gvRaceNone Or CurrCat = CurrentMenu.Category) Then
        Populating = True
        cboView.ListIndex = 0
        For CurrCat = 1 To cboView.ListCount - 1
            If cboView.ItemData(CurrCat) = CurrentMenu.Category Then cboView.ListIndex = CurrCat
        Next CurrCat
        Populating = False
    End If
    
    Set CurrentItem = Nothing
    lvwMenu.ListItems.Clear
    CurrentMenu.First
    Set NewItem = lvwMenu.ListItems.Add(Text:="(back)", Key:="(back)")
    NewItem.Tag = "(back)"
    InsideMenu = True
    lblMenuCaption.Caption = "&Menu: " & CurrentMenu.Name
    cboView.Visible = False
    lblView.Visible = False

    Do Until CurrentMenu.Off
        Set NewItem = lvwMenu.ListItems.Add(Text:=CurrentMenu.DisplayItem, _
                                            Key:="k" & CurrentMenu.ItemName)
        NewItem.Tag = CurrentMenu.ItemName
        If NewItem.Key = SelName Then Found = True
        CurrentMenu.MoveNext
    Loop
    
    If Found Then
        Set lvwMenu.SelectedItem = lvwMenu.ListItems(SelName)
    ElseIf lvwMenu.ListItems.Count >= NextSelIndex Then
        Set lvwMenu.SelectedItem = lvwMenu.ListItems(NextSelIndex)
    ElseIf lvwMenu.ListItems.Count > 0 Then
        Set lvwMenu.SelectedItem = lvwMenu.ListItems(1)
    End If
    If Not lvwMenu.SelectedItem Is Nothing Then lvwMenu.SelectedItem.EnsureVisible
    Call lvwMenu_ItemClick(lvwMenu.SelectedItem)
            
End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Initialize the OutputEngineClass with default output settings.
'
    With OutputEngine
        .Template = tnCharacterSheets
        .GameDate = 0
    End With
    
End Sub

Private Sub cboCategory_Click()
'
' Name:         cboCategory_Click
' Description:  Change the Category type of the current menu.
'

    If Not (Populating Or CurrentMenu Is Nothing) Then
        CurrentMenu.Category = cboCategory.ItemData(cboCategory.ListIndex)
        MenuSet.DataChanged = True
    End If

End Sub

Private Sub cboDisplay_Click()
'
' Name:         cboDisplay_Click
' Description:  Change the display type of the current menu.
'

    If Not (Populating Or CurrentMenu Is Nothing) Then
        CurrentMenu.Display = cboDisplay.ItemData(cboDisplay.ListIndex)
        MenuSet.DataChanged = True
    End If
    
End Sub

Private Sub cboView_Click()
'
' Name:         cboView_Click
' Description:  Refresh the menu list
'

    If Not Populating Then
        RefreshMenus
    End If
    
End Sub

Private Sub chkAddNote_Click()
'
' Name:         chkAddNote_Click
' Description:  Change the add note value on the current menu.
'

    If Not (Populating Or CurrentMenu Is Nothing) Then
        CurrentMenu.Autonote = (chkAddNote.Value = vbChecked)
        MenuSet.DataChanged = True
    End If
    
End Sub

Private Sub chkAlphabetized_Click()
'
' Name:         chkAlphabetized_Click
' Description:  Change the alphabetization on the current menu.  Prompt
'               the user before sorting.
'

    If Not (Populating Or CurrentMenu Is Nothing) Then
        If Not CurrentMenu.IsAlphabetized And chkAlphabetized.Value = vbChecked Then
            If MsgBox("Are you sure you want to sort this menu alphabetically?", _
                        vbYesNo + vbQuestion, "Alphabetize") = vbNo Then
                chkAlphabetized.Value = vbUnchecked
                Exit Sub
            End If
        End If
        CurrentMenu.SetAlphabetized (chkAlphabetized.Value = vbChecked)
        MenuSet.DataChanged = True
    End If
    
End Sub

Private Sub chkNegative_Click()
'
' Name:         chkNegative_Click
' Description:  Change the negative value on the current menu.
'

    If Not (Populating Or CurrentMenu Is Nothing) Then
        CurrentMenu.Negative = (chkNegative.Value = vbChecked)
        MenuSet.DataChanged = True
    End If
    
End Sub

Private Sub chkRequired_Click()
'
' Name:         chkRequired_Click
' Description:  Change the required value on the current menu.
'

    If Not (Populating Or CurrentMenu Is Nothing) Then
        If CurrentMenu.Required And chkRequired.Value = vbUnchecked Then
            If MsgBox("Are you sure you want to remove protection from this required menu?", _
                        vbYesNo + vbQuestion, "Required Menu") = vbNo Then
                chkRequired.Value = vbChecked
                Exit Sub
            End If
        End If
        CurrentMenu.Required = (chkRequired.Value = vbChecked)
        txtMenuName.Locked = CurrentMenu.Required
        MenuSet.DataChanged = True
    End If
    
End Sub

Private Sub cmdCopyMenu_Click()
'
' Name:         cmdCopyMenu_Click
' Description:  Copy the current menu and all its contents.
'

    If Not CurrentMenu Is Nothing Then

        Dim NewMenuName As String
        Dim MenuCopy As LinkedMenuList
        
        NewMenuName = CreateNewName(MenuSet, CurrentMenu.Name & " Copy")
        
        Set MenuCopy = MenuSet.AddNewMenu(NewMenuName)
        MenuCopy.Category = CurrentMenu.Category
        MenuCopy.SetAlphabetized CurrentMenu.IsAlphabetized
        MenuCopy.Autonote = CurrentMenu.Autonote
        MenuCopy.Negative = CurrentMenu.Negative
        MenuCopy.Required = CurrentMenu.Required
        MenuCopy.Display = CurrentMenu.Display
        
        CurrentMenu.First
        Do Until CurrentMenu.Off
            MenuCopy.Append CurrentMenu.ItemName, CurrentMenu.ItemCost, CurrentMenu.ItemNote
            CurrentMenu.MoveNext
        Loop
        
        RefreshMenus MenuCopy
        tabTabs.Tabs("Items").Selected = True
        txtMenuName.SetFocus
        MenuSet.DataChanged = True

    End If

End Sub

Private Sub cmdDelete_Click()
'
' Name:         cmdDelete_Click
' Description:  Delete the currently selected item.
'

    Dim Warning As String
    Dim NextItem As ListItem
    Dim Continue As Boolean
    
    If MenuSet.IsEmpty Or CurrentMenu Is Nothing Or CurrentItem Is Nothing Then Exit Sub
    If InsideMenu And CurrentItem.Tag = "(back)" Then Exit Sub
    
    If (Not InsideMenu) And CurrentMenu.Required Then
        MsgBox "This menu is required by Grapevine and should not be deleted.", _
                vbOKOnly + vbExclamation, "Required Menu"
    Else

        If chkConfirmDelete.Value = vbChecked Then
        
            If Not InsideMenu Then
                Warning = "Are you sure you want to delete the """ & CurrentMenu.Name _
                        & """ menu and all its contents?"
                Continue = MsgBox(Warning, vbYesNo + vbQuestion, "Delete") = vbYes
            Else
                Warning = "Are you sure you want to delete """ & CurrentItem.Tag & """?"
                Continue = MsgBox(Warning, vbYesNo + vbQuestion, "Delete") = vbYes
            End If
            
        Else
            Continue = True
        End If
        
        If Continue Then
        
            If CurrentItem.Index < lvwMenu.ListItems.Count Then
                Set NextItem = lvwMenu.ListItems(CurrentItem.Index + 1)
            ElseIf CurrentItem.Index > 1 Then
                Set NextItem = lvwMenu.ListItems(CurrentItem.Index - 1)
            End If
            
            If Not InsideMenu Then
                MenuSet.RemoveMenu CurrentMenu.Name
            Else
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.Remove
            End If
            MenuSet.DataChanged = True

            lvwMenu.ListItems.Remove CurrentItem.Index
            Set lvwMenu.SelectedItem = NextItem
            Call lvwMenu_ItemClick(NextItem)
            
        End If

    End If

    lvwMenu.SetFocus

End Sub

Private Sub cmdESCClose_Click()
'
' Name:         cmdESCClose_Click
' Description:  Close the window when the user presses ESC.
'

    Unload Me

End Sub

Private Sub cmdFind_Click()
'
' Name:         cmdFind_Click
' Description:  Find the next menu item matching the criteria.
'

    Dim FirstMenuFound As LinkedMenuList
    Dim FirstItemFound As String
    Dim aMenu As LinkedMenuList
    Dim Found As Boolean
    Dim FoundItem As String
    
    If Not (txtFind.Text = "" Or CurrentMenu Is Nothing Or CurrentItem Is Nothing) Then
    
        Screen.MousePointer = vbHourglass
    
        If InsideMenu Then
            CurrentMenu.MoveTo CurrentItem.Tag
            CurrentMenu.MoveNext
        Else
            CurrentMenu.First
        End If
        Do Until CurrentMenu.Off
            If optExact.Value Then
                Found = LCase(txtFind.Text) = LCase(CurrentMenu.ItemName)
            Else
                Found = InStr(LCase(CurrentMenu.ItemName), LCase(txtFind.Text)) > 0
            End If
            If Found Then
                RefreshMenuItems CurrentMenu, CurrentMenu.ItemName
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            CurrentMenu.MoveNext
        Loop
    
        MenuSet.First
        Do Until MenuSet.Off
            
            Set aMenu = MenuSet.Menu
            If optExact.Value Then
                Found = LCase(txtFind.Text) = LCase(aMenu.Name)
            Else
                Found = InStr(LCase(aMenu.Name), LCase(txtFind.Text)) > 0
            End If
            If Not Found Then
                aMenu.First
                Do Until aMenu.Off
                    If optExact.Value Then
                        Found = LCase(txtFind.Text) = LCase(aMenu.ItemName)
                    Else
                        Found = InStr(LCase(aMenu.ItemName), LCase(txtFind.Text)) > 0
                    End If
                    If Found Then
                        FoundItem = aMenu.ItemName
                        Exit Do
                    End If
                    aMenu.MoveNext
                Loop
            End If
            
            If Found Then
                If aMenu.Name > CurrentMenu.Name Then
                    If FoundItem = "" Then
                        RefreshMenus aMenu
                    Else
                        RefreshMenuItems aMenu, FoundItem
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf FirstMenuFound Is Nothing Then
                    Set FirstMenuFound = aMenu
                    FirstItemFound = FoundItem
                    Found = False
                    FoundItem = ""
                End If
            End If
            
            MenuSet.MoveNext
        Loop
    
        If Not FirstMenuFound Is Nothing Then
            If FirstItemFound = "" Then
                RefreshMenus FirstMenuFound
            Else
                RefreshMenuItems FirstMenuFound, FirstItemFound
            End If
        Else
            MsgBox "Search text is not found.", vbOKOnly + vbExclamation, "Find"
        End If
    
        Screen.MousePointer = vbDefault
    
    End If

End Sub

Private Sub cmdFivePower_Click()
'
' Name:         cmdFivePower_Click
' Description:  Create a new menu of five powers.
'

    Dim NewMenuName As String
    Dim NewMenuNode As Node
    Dim FivePowerMenu As LinkedMenuList
    
    NewMenuName = CreateNewName(MenuSet, "New Power")
    
    Set FivePowerMenu = MenuSet.AddNewMenu(NewMenuName)
    FivePowerMenu.Category = cboView.ItemData(cboView.ListIndex)
    If FivePowerMenu.Category = gvRaceNone Then FivePowerMenu.Category = gvRaceAll
    FivePowerMenu.SetAlphabetized False
    FivePowerMenu.Autonote = True
    FivePowerMenu.Display = ldNoteOnly
    FivePowerMenu.Append "First Basic", "3", "basic"
    FivePowerMenu.Append "Second Basic", "3", "basic"
    FivePowerMenu.Append "First Intermediate", "6", "int."
    FivePowerMenu.Append "Second Intermediate", "6", "int."
    FivePowerMenu.Append "Advanced", "9", "adv."

    RefreshMenus FivePowerMenu
    
    tabTabs.Tabs("Items").Selected = True
    txtMenuName.SetFocus
    MenuSet.DataChanged = True

End Sub

Private Sub cmdGoToIncludeMenu_Click()
'
' Name:         cmdGoToIncludeMenu_Click
' Description:  Jump to the included menu.
'

    Dim FindMenu As LinkedMenuList
    Set FindMenu = MenuSet.GetMenu(lstIncludeMenu.Text)
    If Not FindMenu Is Nothing Then RefreshMenuItems FindMenu, "1"
    
End Sub

Private Sub cmdGoToSubmenu_Click()
'
' Name:         cmdGoToSubmenu_Click
' Description:  Jump to the linked submenu.
'

    Dim FindMenu As LinkedMenuList
    Set FindMenu = MenuSet.GetMenu(lstLinkedMenu.Text)
    If Not FindMenu Is Nothing Then RefreshMenuItems FindMenu, "1"

End Sub

Private Sub cmdIncludeMenu_Click()
'
' Name:         cmdIncludeMenu_Click
' Description:  Create a new include menu item under the current menu.
'               If a menu is alphabetized, it is inserted in sorted
'               position.  If an ordered menu is selected, it is
'               the first item.  If any other type of choice in an
'               ordered menu is selected, it is inserted before the
'               selection.
'

    If Not CurrentMenu Is Nothing Then
        CreateNewMenuEntry "New Include", "+", ""
        txtIncludeName.SetFocus
    End If

End Sub

Private Sub cmdLoadMenu_Click()
'
' Name:         cmdLoadMenu_Click
' Description:  Prompt the user for a filename, then Load a new Menu.
'

    Dim Continue As Boolean
    Dim aMenu As LinkedMenuList
    Dim NewMenuNode As Node
    Dim NodeText As String
    
    PromptForSave Continue
    
    If Continue Then
        cmnDialog.DialogTitle = "Load Menu"
        cmnDialog.InitDir = GetSetting(App.Title, "Files", "MenuDir", CurDir)
        cmnDialog.FileName = ""
        cmnDialog.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
        cmnDialog.Filter = "Grapevine Menu Files (*.gvm)|*.gvm|All Files|*.*"
        cmnDialog.FilterIndex = 1
        
        On Error GoTo LoadMenu_AnyError
        cmnDialog.ShowOpen
        On Error GoTo 0
    
        SaveSetting App.Title, "Files", "MenuDir", CurDir
    
        Screen.MousePointer = vbHourglass
        mdiMain.pgbProgress.Value = 0
        mdiMain.pgbProgress.Max = 101
        mdiMain.pgbProgress.Visible = True
        
        MenuSet.OpenMenus cmnDialog.FileName, True
        
        If MenuSet.FileError Then
            MsgBox MenuSet.FileErrorMessage, vbExclamation, "Load Menus"
        End If
            
        Populating = True
        lblFilename.Caption = MenuSet.FileName
        txtPath.Text = MenuSet.FilePath
        txtDescription.Text = MenuSet.Description
        Populating = False
        
        RefreshMenus
        
        mdiMain.pgbProgress.Visible = False
        Screen.MousePointer = vbDefault
    
    End If
    
    GoTo LoadMenu_Finish
    
LoadMenu_AnyError:
    Resume LoadMenu_Finish
LoadMenu_Finish:

End Sub

Private Sub cmdMergeMenu_Click()
'
' Name:         cmdMergeMenu_Click
' Description:  Load a second menu and merge it with this one.
'

    Dim Merger As Integer
    Dim Prompt As String
    Dim MergeFile As String
    
    With cmnDialog
        .DialogTitle = "Load Menu"
        .InitDir = GetSetting(App.Title, "Files", "MenuDir", CurDir)
        .FileName = ""
        .Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
        .Filter = "Grapevine Menu Files and Updates (*.gvm;*.gvu)|*.gvm;*.gvu|All Files|*.*"
        .FilterIndex = 1
        On Error GoTo MergeMenu_AnyError
        .ShowOpen
        On Error GoTo 0
    End With
    
    MergeFile = cmnDialog.FileName
    
    Prompt = "Click YES to do an Aggressive Merge, in which differences between the two" & vbCrLf & _
             "menu files are resolved in favor of the file you are loading (" & ShortFile(MergeFile) & ")." & _
             vbCrLf & "This is recommended if you have not made many changes to your menus." & vbCrLf & vbCrLf
    Prompt = Prompt & "Click NO to do a Conservative Merge, which resolves differences in favor of" & vbCrLf & _
             "your existing menu file (" & ShortFile(MenuSet.FileName) & "). This is recommended if" & vbCrLf & _
             "you have made extensive changes to your menus." & vbCrLf & vbCrLf
    Prompt = Prompt & "All changes made will be described in the file MergeLog.txt."
             
    Merger = MsgBox(Prompt, vbYesNoCancel + vbQuestion, "Merge Menus")
    
    If Merger <> vbCancel Then
        
        Screen.MousePointer = vbHourglass
        mdiMain.pgbProgress.Value = 0
        mdiMain.pgbProgress.Max = 201
        mdiMain.pgbProgress.Visible = True
        
        MenuSet.MergeMenus MergeFile, (Merger = vbYes)
        
        If MenuSet.FileError Then
            MsgBox MenuSet.FileErrorMessage, vbExclamation, "Merge Menus"
        End If
            
        txtDescription.Text = MenuSet.Description
        MenuSet.DataChanged = True
        
        RefreshMenus
        
        mdiMain.pgbProgress.Visible = False
        Screen.MousePointer = vbDefault
    
        Prompt = "Would you like to view the changes recorded in the file MergeLog.txt?"
        Merger = MsgBox(Prompt, vbYesNo + vbQuestion, "View Log")
        If Merger = vbYes Then
            ShellExecute mdiMain.hWnd, "open", SlashPath(App.Path) & "MergeLog.txt", "", "", 1
        End If
        
    End If

    GoTo MergeMenu_Finish
MergeMenu_AnyError:
    Resume MergeMenu_Finish
MergeMenu_Finish:

End Sub

Private Sub cmdNewMenu_Click()
'
' Name:         cmdNewMenu_Click
' Description:  Create a new menu.
'

    Dim NewMenuName As String
    Dim NewMenuNode As Node
    Dim NewMenu As LinkedMenuList
    
    NewMenuName = CreateNewName(MenuSet, "New Menu")
    
    Set NewMenu = MenuSet.AddNewMenu(NewMenuName)
    If Not cboView.ItemData(cboView.ListIndex) = gvRaceNone Then
        NewMenu.Category = cboView.ItemData(cboView.ListIndex)
    End If
    RefreshMenus NewMenu
    
    txtMenuName.SetFocus
    MenuSet.DataChanged = True

End Sub

Private Sub cmdNewMenuItem_Click()
'
' Name:         cmdNewMenuItem_Click
' Description:  Create a new menu item under the current menu.  If
'               a menu is alphabetized, it is inserted in sorted
'               position.  If an ordered menu is selected, it is
'               the first item.  If any other type of choice in an
'               ordered menu is selected, it is inserted before the
'               selection.
'

    If Not CurrentMenu Is Nothing Then
        CreateNewMenuEntry "New Item", "1", ""
        txtItemName.SetFocus
    End If
    
End Sub

Private Sub cmdNewSubmenu_Click()
'
' Name:         cmdNewSubmenu_Click
' Description:  Create a new submenu link under the current menu.  If
'               a menu is alphabetized, it is inserted in sorted
'               position.  If an ordered menu is selected, it is
'               the first item.  If any other type of choice in an
'               ordered menu is selected, it is inserted before the
'               selection.
'

    If Not CurrentMenu Is Nothing Then
        CreateNewMenuEntry "New Submenu", ":", ""
        txtSubmenuName.SetFocus
    End If
    
End Sub

Private Sub cmdSaveMenu_Click()
'
' Name:         mnuSaveMenu_Click
' Description:  Save the current Menu, prompting if there is no filename yet.
'

    If MenuSet.FilePath = "" Then
        Call cmdSaveMenuAs_Click
    Else
        ValidateControls
        Screen.MousePointer = vbHourglass
        mdiMain.pgbProgress.Value = 0
        mdiMain.pgbProgress.Max = 101
        mdiMain.pgbProgress.Visible = True
        MenuSet.SaveMenus MenuSet.FilePath
        mdiMain.pgbProgress.Visible = False
        Screen.MousePointer = vbDefault
        If MenuSet.FileError Then _
            MsgBox MenuSet.FileErrorMessage, vbExclamation, "Save Menus"
    End If
    

End Sub

Private Sub cmdSaveMenuAs_Click()
'
' Name:         mnuSaveMenuAs_Click
' Description:  Prompt for a filename and save the curent Menu.
'

    ValidateControls

    cmnDialog.DialogTitle = "Save Menu As..."
    cmnDialog.InitDir = GetSetting(App.Title, "Files", "MenuDir", CurDir)
    cmnDialog.FileName = MenuSet.FileName
    cmnDialog.Flags = cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + _
            cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    cmnDialog.Filter = "Grapevine Menu File (Binary Format)|*.gvm|" & _
                       "Grapevine Menu File (XML Format)|*.gvm|" & _
                       "All Files|*.*"
    cmnDialog.FilterIndex = IIf(MenuSet.FileFormat = gvXML, 2, 1)
    
    On Error GoTo mnuSaveMenuAs_AnyError
    cmnDialog.ShowSave
    On Error GoTo 0

    SaveSetting App.Title, "Files", "MenuDir", CurDir
    MenuSet.FileFormat = IIf(cmnDialog.FilterIndex = 1, gvBinaryMenu, gvXML)
    
    Screen.MousePointer = vbHourglass
    mdiMain.pgbProgress.Value = 0
    mdiMain.pgbProgress.Max = 101
    mdiMain.pgbProgress.Visible = True
    MenuSet.SaveMenus cmnDialog.FileName
    mdiMain.pgbProgress.Visible = False
    Screen.MousePointer = vbDefault
    
    If MenuSet.FileError Then
        MsgBox MenuSet.FileErrorMessage, vbExclamation, "Save Menu"
    Else
        lblFilename.Caption = MenuSet.FileName
        txtPath.Text = MenuSet.FilePath
    End If
    
    GoTo mnuSaveMenuAs_Finish
    
mnuSaveMenuAs_AnyError:
    Resume mnuSaveMenuAs_Finish
mnuSaveMenuAs_Finish:
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
' Name:         Form_KeyDown
' Description:  Common keyboard shortcuts.
'
    Select Case KeyCode
        Case vbKeyPageUp
            SendKeys "%M"
    End Select
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the menu editor window.
'

    Dim aMenu As LinkedMenuList
    Dim NewMenuNode As Node
    Dim NodeText As String
    Dim I As Integer
    
    Set MenuSet = Game.MenuSet
    
    Populating = True
    Screen.MousePointer = vbHourglass

    lblFilename.Caption = MenuSet.FileName
    txtPath.Text = MenuSet.FilePath
    txtDescription.Text = MenuSet.Description
    
    cboCategory.AddItem "General", 0
    cboCategory.ItemData(0) = gvRaceAll
    cboCategory.AddItem "Changeling", 1
    cboCategory.ItemData(1) = gvRaceChangeling
    cboCategory.AddItem "Demon", 2
    cboCategory.ItemData(2) = gvRaceDemon
    cboCategory.AddItem "Fera", 3
    cboCategory.ItemData(3) = gvRaceFera
    cboCategory.AddItem "Hunter", 4
    cboCategory.ItemData(4) = gvRaceHunter
    cboCategory.AddItem "Kuei-Jin", 5
    cboCategory.ItemData(5) = gvRaceKueiJin
    cboCategory.AddItem "Mage", 6
    cboCategory.ItemData(6) = gvracemage
    cboCategory.AddItem "Mortal", 7
    cboCategory.ItemData(7) = gvRaceMortal
    cboCategory.AddItem "Mummy", 8
    cboCategory.ItemData(8) = gvRaceMummy
    cboCategory.AddItem "Vampire", 9
    cboCategory.ItemData(9) = gvRaceVampire
    cboCategory.AddItem "Various", 10
    cboCategory.ItemData(10) = gvRaceVarious
    cboCategory.AddItem "Werewolf", 11
    cboCategory.ItemData(11) = gvRaceWerewolf
    cboCategory.AddItem "Wraith", 12
    cboCategory.ItemData(12) = gvRaceWraith
    
    cboView.AddItem "All Menus", 0
    cboView.ItemData(0) = gvRaceNone
    For I = 0 To cboCategory.ListCount - 1
        cboView.AddItem cboCategory.List(I) & " Menus", I + 1
        cboView.ItemData(I + 1) = cboCategory.ItemData(I)
    Next I
    
    cboDisplay.AddItem "Simple", 0
    cboDisplay.ItemData(0) = ldSimple
    cboDisplay.AddItem "Cost and Note", 1
    cboDisplay.ItemData(1) = ldCost
    cboDisplay.AddItem "Cost Only", 2
    cboDisplay.ItemData(2) = ldCostOnly
    cboDisplay.AddItem "Note Only", 3
    cboDisplay.ItemData(3) = ldNoteOnly
    
    CurrentTabKey = ""
    tabTabs.Tabs("Items").Selected = True
    cboView.ListIndex = 1
    
    RefreshMenus
    
    mdiMain.OrientForm Me
    Me.Show
    lvwMenu.SetFocus

    Screen.MousePointer = vbDefault
    Populating = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' Name:         Form_QueryUnload
' Parameters:   Cancel          whether or not to untlimately unload
'               UnloadMode      the reason this form is being unloaded
' Description:  When the user closes the window, ensure he wants to
'               save his changes.
'

    Dim Continue As Boolean
    PromptForSave Continue
    Cancel = Not Continue
    If Not Cancel Then MenuSet.DataChanged = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Name:         Form_Unload
' Description:  Remove this form from memory

    Me.Hide
    Unload Me

End Sub

Private Sub lstIncludeMenu_DblClick()
'
' Name:         lstIncludeMenu_DblClick
' Description:  Jump to the included menu.
'
    ValidateControls
    Call cmdGoToIncludeMenu_Click

End Sub

Private Sub lstIncludeMenu_Validate(Cancel As Boolean)
'
' Name:         lstIncludeMenu_Validate
' Parameters:   Cancel          more of a return value -- set TRUE to cancel validation
' Description:  Change the included menu of this item.
'

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing) Then
        CurrentMenu.MoveTo CurrentItem.Tag
        If Not CurrentMenu.Off Then
            CurrentItem.Text = CurrentMenu.DisplayItem
            CurrentMenu.SetItemNote lstIncludeMenu.Text
            MenuSet.DataChanged = True
        End If
    End If
    
End Sub

Private Sub lstLinkedMenu_DblClick()
'
' Name:         lstLinkedMenu_DblClick
' Description:  Jump to the linked submenu.
'
    ValidateControls
    Call cmdGoToSubmenu_Click

End Sub

Private Sub lstLinkedMenu_Validate(Cancel As Boolean)
'
' Name:         lstLinkedMenu_Validate
' Parameters:   Cancel          mare of a return value -- set TRUE to cancel validation
' Description:  Change the linked menu of this submenu.
'
    
    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing) Then
        CurrentMenu.MoveTo CurrentItem.Tag
        If Not CurrentMenu.Off Then
            CurrentItem.Text = CurrentMenu.DisplayItem
            CurrentMenu.SetItemNote lstLinkedMenu.Text
            MenuSet.DataChanged = True
        End If
    End If
    
End Sub

Private Sub lvwMenu_DblClick()
'
' Name:         lvwMenu_DblClick
' Description:  Go to a submenu, if this is a menu being clicked
'
    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing) Then
        If Not InsideMenu Then
            RefreshMenuItems CurrentMenu, "1"
        ElseIf CurrentItem.Tag = "(back)" Then
            RefreshMenus CurrentMenu
        ElseIf fraIncludeMenu.Visible Then
            Call cmdGoToIncludeMenu_Click
        ElseIf fraSubmenu.Visible Then
            Call cmdGoToSubmenu_Click
        End If
    End If
    
End Sub

Private Sub lvwMenu_KeyPress(KeyAscii As Integer)
'
' Name:         lvwMenu_KeyDown
' Parameters:   KeyPress        the key pressed
' Description:  Implement keyboard shortcuts for common functions.

    Select Case KeyAscii
        Case Asc("+"), Asc("=")
            If updPosition.Visible Then updPosition_UpClick
            KeyAscii = 0
        Case Asc("-")
            If updPosition.Visible Then updPosition_DownClick
            KeyAscii = 0
    End Select
    
End Sub

Private Sub tabTabs_Click()
'
' Name:         tabTabs_Click
' Description:  Show the appropriate frame when a tab is clicked.
'
    
    If CurrentTabKey <> tabTabs.SelectedItem.Key Then
    
        CurrentTabKey = tabTabs.SelectedItem.Key
        Select Case CurrentTabKey
            Case "File"
                fraFileFrame.Visible = True
                fraItemFrame.Visible = False
                fraToolFrame.Visible = False
            Case "Items"
                fraItemFrame.Visible = True
                fraFileFrame.Visible = False
                fraToolFrame.Visible = False
            Case "Tools"
                fraToolFrame.Visible = True
                fraFileFrame.Visible = False
                fraItemFrame.Visible = False
        End Select
        
    End If

End Sub

Private Sub lvwMenu_KeyDown(KeyCode As Integer, Shift As Integer)
'
' Name:         lvwMenu_KeyDown
' Parameters:   KeyCode         the key pressed
'               Shift           state of Shift, Alt, Ctrl
' Description:  Implement keyboard shortcuts for common functions.
'

    Select Case KeyCode
        Case vbKeyReturn
            If fraItemFrame.Visible Then
                If fraMenu.Visible Then txtMenuName.SetFocus
                If fraSubmenu.Visible Then txtSubmenuName.SetFocus
                If fraIncludeMenu.Visible Then txtIncludeName.SetFocus
                If fraItem.Visible Then txtItemName.SetFocus
            End If
            KeyCode = 0
        Case vbKeyInsert
            If fraItemFrame.Visible Then cmdNewMenuItem_Click
            KeyCode = 0
        Case vbKeyBack
            If InsideMenu And Not CurrentMenu Is Nothing Then RefreshMenus CurrentMenu
            KeyCode = 0
        Case vbKeyDelete
            If fraItemFrame.Visible Then cmdDelete_Click
            KeyCode = 0
        Case vbKeySpace
            Call lvwMenu_DblClick
            KeyCode = 0
    End Select

    '
    ' Special secret shortcuts
    '
    If (Shift And vbCtrlMask) > 0 And Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing) Then
        
        Dim NoItemRefresh As Boolean
        
        Select Case KeyCode

            Case vbKeyB
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "3"
                CurrentMenu.SetItemNote "basic"
            Case vbKeyI
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "6"
                CurrentMenu.SetItemNote "int."
            Case vbKeyA
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "9"
                CurrentMenu.SetItemNote "adv."
            Case vbKeyE
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "12"
                CurrentMenu.SetItemNote "elder"
            Case vbKeyM
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "15"
                CurrentMenu.SetItemNote "master"
            Case vbKeyS
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "18"
                CurrentMenu.SetItemNote "asc."
            Case vbKeyT
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "21"
                CurrentMenu.SetItemNote "meth."
            Case vbKeyN
                CurrentMenu.MoveTo CurrentItem.Tag
                CurrentMenu.SetItemCost "12"
                CurrentMenu.SetItemNote "master"
            Case vbKeyR
                CurrentMenu.MoveTo CurrentItem.Tag
                Select Case CurrentMenu.ItemCost
                    Case "3": CurrentMenu.SetItemCost "2"
                    Case "6": CurrentMenu.SetItemCost "4"
                    Case "9": CurrentMenu.SetItemCost "7"
                    Case "12": CurrentMenu.SetItemCost "10"
                End Select
                CurrentMenu.SetItemNote (CurrentMenu.ItemNote & " ritual")
            Case vbKeyV
                CurrentMenu.MoveTo CurrentItem.Tag
                Select Case CurrentMenu.ItemCost
                    Case "3": CurrentMenu.SetItemCost "2"
                    Case "6": CurrentMenu.SetItemCost "4"
                    Case "9": CurrentMenu.SetItemCost "6"
                End Select
            Case vbKeyG
                CurrentMenu.MoveTo CurrentItem.Tag
                Select Case CurrentMenu.ItemCost
                    Case "3": CurrentMenu.SetItemCost "2"
                    Case "6": CurrentMenu.SetItemCost "4"
                    Case "9": CurrentMenu.SetItemCost "6"
                End Select
            Case Else
                NoItemRefresh = True
        End Select
        
        If Not NoItemRefresh Then
            If Not CurrentMenu.Off Then
                CurrentItem.Text = CurrentMenu.DisplayItem
                txtItemName.Text = CurrentMenu.ItemName
                txtItemCost.Text = CurrentMenu.ItemCost
                txtItemNote.Text = CurrentMenu.ItemNote
                KeyCode = 0
            End If
        End If
        
    End If
        
End Sub

Private Sub lvwMenu_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwMenu_ItemClick
' Parameters:   Item            the node clicked
' Description:  Display the frame appropriate to the node, populated
'               with information from the menu or menu item
'

    If Not Item Is Nothing Then
    
        Populating = True
    
        fraItem.Visible = False
        fraSubmenu.Visible = False
        fraIncludeMenu.Visible = False
        fraMenu.Visible = False
        
        Set CurrentItem = Item
                
        If InsideMenu And Not CurrentMenu Is Nothing Then
            
            If Item.Key <> "(back)" Then

                lblPosition.Visible = Not CurrentMenu.IsAlphabetized
                updPosition.Visible = Not CurrentMenu.IsAlphabetized
                CurrentMenu.MoveTo Item.Tag
                If Not CurrentMenu.Off Then
                
                    Select Case CurrentMenu.ItemCost
                        Case ":"                    'this is a submenu link
                            txtSubmenuName.Text = CurrentMenu.ItemName
                            lstLinkedMenu.Clear
                            MenuSet.First
                            Do Until MenuSet.Off
                                lstLinkedMenu.AddItem MenuSet.Menu.Name
                                If MenuSet.Menu.Name = CurrentMenu.ItemNote Then
                                    lstLinkedMenu.ListIndex = lstLinkedMenu.NewIndex
                                End If
                                MenuSet.MoveNext
                            Loop
                            If lstLinkedMenu.ListIndex > -1 Then _
                                lstLinkedMenu.TopIndex = lstLinkedMenu.ListIndex
                            fraSubmenu.Visible = True
                        Case "+"                    'this is an include
                            txtIncludeName.Text = CurrentMenu.ItemName
                            lstIncludeMenu.Clear
                            MenuSet.First
                            Do Until MenuSet.Off
                                lstIncludeMenu.AddItem MenuSet.Menu.Name
                                If MenuSet.Menu.Name = CurrentMenu.ItemNote Then
                                    lstIncludeMenu.ListIndex = lstIncludeMenu.NewIndex
                                End If
                                MenuSet.MoveNext
                            Loop
                            If lstIncludeMenu.ListIndex > -1 Then _
                                lstIncludeMenu.TopIndex = lstIncludeMenu.ListIndex
                            fraIncludeMenu.Visible = True
                        Case Else                   'this is an ordinary item
                            txtItemName.Text = CurrentMenu.ItemName
                            txtItemCost.Text = CurrentMenu.ItemCost
                            txtItemNote.Text = CurrentMenu.ItemNote
                            fraItem.Visible = True
                    End Select
                
                End If
            
            Else
                lblPosition.Visible = False
                updPosition.Visible = False
            End If
            
        ElseIf Not InsideMenu Then

            Dim FindMenu As LinkedMenuList
            Dim I As Integer
    
            Set FindMenu = MenuSet.GetMenu(Item.Tag)
            If Not FindMenu Is Nothing Then
                
                Set CurrentMenu = FindMenu
                txtMenuName.Text = CurrentMenu.Name
                txtMenuName.Locked = CurrentMenu.Required
                chkAlphabetized.Value = IIf(CurrentMenu.IsAlphabetized, vbChecked, vbUnchecked)
                chkNegative.Value = IIf(CurrentMenu.Negative, vbChecked, vbUnchecked)
                chkAddNote.Value = IIf(CurrentMenu.Autonote, vbChecked, vbUnchecked)
                chkRequired.Value = IIf(CurrentMenu.Required, vbChecked, vbUnchecked)
                
                For I = 0 To cboCategory.ListCount - 1
                    If cboCategory.ItemData(I) = CurrentMenu.Category Then
                        cboCategory.ListIndex = I
                        Exit For
                    End If
                Next I
                
                For I = 0 To cboDisplay.ListCount - 1
                    If cboDisplay.ItemData(I) = CurrentMenu.Display Then
                        cboDisplay.ListIndex = I
                        Exit For
                    End If
                Next I
                
                fraMenu.Visible = True
            
            Else
                
                Set CurrentMenu = Nothing
            
            End If
            
            lblPosition.Visible = False
            updPosition.Visible = False
        
        End If

        Populating = False

    Else
        
        If Not InsideMenu Then Set CurrentMenu = Nothing
        Set CurrentItem = Nothing
        
    End If
    
End Sub

Private Sub txtDescription_GotFocus()
'
' Name:         txtDescription_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtDescription

End Sub

Private Sub txtDescription_Validate(Cancel As Boolean)
'
' Name:         txtDescription_Validate
' Parameters:   Cancel          whether to stop validation
' Description:  Store the description.
'

    If Not Populating Then
        MenuSet.DataChanged = MenuSet.DataChanged Or _
                              Not (MenuSet.Description = txtDescription.Text)
        MenuSet.Description = txtDescription.Text
    End If
    
End Sub

Private Sub txtFind_GotFocus()
'
' Name:         txtFind_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtFind

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
'
' Name:         txtFind_KeyPress
' Parameters:   KeyAscii            the key pressed
' Description:  Find the item when return is pressed.
'
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
        KeyAscii = 0
    End If

End Sub

Private Sub txtItemCost_GotFocus()
'
' Name:         txtItemCost_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtItemCost

End Sub

Private Sub txtItemCost_KeyPress(KeyAscii As Integer)
'
' Name:         txtItemCost_KeyPress
' Parameters:   KeyAscii            the key pressed
' Description:  Move focus to the next control when return is pressed.
'
    Dim Cancel As Boolean
    If KeyAscii = 13 Then
        Call txtItemCost_Validate(Cancel)
        If Not Cancel Then txtItemNote.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtItemCost_Validate(Cancel As Boolean)
'
' Name:         txtItemCost_Validate
' Parameters:   Cancel          more a return value -- set TRUE to cancel validation
' Description:  Ensure the given cost is legal before changing it.
'

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing Or Populating) Then
        txtItemCost.Text = Trim(txtItemCost.Text)
        CurrentMenu.MoveTo CurrentItem.Tag
        If Not CurrentMenu.Off Then
            If Not txtItemCost.Text = CurrentMenu.ItemCost Then
                If txtItemCost.Text = ":" Or txtItemCost.Text = "+" Then
                    MsgBox "Colon or plus-sign characters cannot be used as a cost.", _
                            vbOKOnly + vbExclamation, "Bad Text"
                    txtItemCost.Text = CurrentMenu.ItemCost
                    txtItemCost.SelStart = 0
                    txtItemCost.SelLength = Len(txtItemCost.Text)
                    Cancel = True
                Else
                    CurrentMenu.SetItemCost txtItemCost.Text
                    CurrentItem.Text = CurrentMenu.DisplayItem
                    MenuSet.DataChanged = True
                End If
            End If
        End If
    End If
    
End Sub

Private Sub txtItemName_GotFocus()
'
' Name:         txtItemName_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtItemName

End Sub

Private Sub txtItemName_KeyPress(KeyAscii As Integer)
'
' Name:         txtItemName_KeyPress
' Parameters:   KeyAscii            the key pressed
' Description:  Move focus to the next control when return is pressed.
'
   
    Dim Cancel As Boolean
    If KeyAscii = vbKeyReturn Then
        Call txtItemName_Validate(Cancel)
        If Not Cancel Then txtItemCost.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtItemName_Validate(Cancel As Boolean)
'
' Name:         txtItemName_Validate
' Parameters:   Cancel          more a return value -- set TRUE to cancel validation
' Description:  Ensure the given name is legal before changing it.
'

    Dim CaseChange As Boolean

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing Or Populating) Then
        With txtItemName
            
            .Text = Trim(.Text)
            If Not (.Text = CurrentItem.Tag Or .Text = "") Then
                
                Cancel = False
                CaseChange = (LCase(.Text) = LCase(CurrentItem.Tag))
                If Not CaseChange Then
                    Cancel = Not ValidateName(txtItemName, CurrentMenu, CurrentItem.Tag)
                End If
                If Not Cancel Then
                    CurrentMenu.SetItemName CurrentItem.Tag, .Text
                    MenuSet.DataChanged = True
                    If CaseChange Then
                        CurrentItem.Tag = .Text
                        CurrentItem.Text = CurrentMenu.DisplayItem
                    Else
                        RefreshMenuItems CurrentMenu, .Text
                    End If
                End If
                
            Else
                .Text = CurrentItem.Tag
            End If
        
        End With
    End If

End Sub

Private Sub txtItemNote_GotFocus()
'
' Name:         txtItemNote_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtItemNote

End Sub

Private Sub txtItemNote_KeyPress(KeyAscii As Integer)
'
' Name:         txtItemNote_KeyPress
' Parameters:   KeyAscii            the key pressed
' Description:  Move focus to the next control when return is pressed.
'
   
    Dim Cancel As Boolean
    If KeyAscii = 13 Then
        Call txtItemNote_Validate(Cancel)
        If Not Cancel Then cmdNewMenuItem.SetFocus
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtItemNote_Validate(Cancel As Boolean)
'
' Name:         txtItemNote_Validate
' Parameters:   Cancel          more a return value -- set TRUE to cancel validation
' Description:  Change the note on the item.
'

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing Or Populating) Then
        txtItemNote.Text = Trim(txtItemNote.Text)
        CurrentMenu.MoveTo CurrentItem.Tag
        If Not CurrentMenu.Off Then
            If Not txtItemNote.Text = CurrentMenu.ItemNote Then
                CurrentMenu.SetItemNote txtItemNote.Text
                CurrentItem.Text = CurrentMenu.DisplayItem
                MenuSet.DataChanged = True
            End If
        End If
    End If
    
End Sub

Private Sub txtIncludeName_GotFocus()
'
' Name:         txtIncludeName_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtIncludeName

End Sub

Private Sub txtIncludeName_KeyPress(KeyAscii As Integer)
'
' Name:         txtIncludeName_KeyPress
' Parameters:   KeyAscii            the key pressed
' Description:  Move focus to the next control when return is pressed.
'
   
    Dim Cancel As Boolean
    If KeyAscii = 13 Then
        Call txtIncludeName_Validate(Cancel)
        If Not Cancel Then lstIncludeMenu.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtIncludeName_Validate(Cancel As Boolean)
'
' Name:         txtIncludeName_Validate
' Parameters:   Cancel          more a return value -- set TRUE to cancel validation
' Description:  Ensure the given name is legal before changing it.
'

    Dim CaseChange As Boolean

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing Or Populating) Then
        With txtIncludeName
            
            .Text = Trim(.Text)
            If Not (.Text = CurrentItem.Tag Or .Text = "") Then
                
                Cancel = False
                CaseChange = (LCase(.Text) = LCase(CurrentItem.Tag))
                If Not CaseChange Then
                    Cancel = Not ValidateName(txtIncludeName, CurrentMenu, CurrentItem.Tag)
                End If
                If Not Cancel Then
                    CurrentMenu.SetItemName CurrentItem.Tag, .Text
                    MenuSet.DataChanged = True
                    If CaseChange Then
                        CurrentItem.Tag = .Text
                        CurrentItem.Text = CurrentMenu.DisplayItem
                    Else
                        RefreshMenuItems CurrentMenu, .Text
                    End If
                End If
                
            Else
                .Text = CurrentItem.Tag
            End If
        
        End With
    End If
    
End Sub

Private Sub txtMenuName_GotFocus()
'
' Name:         txtMenuName_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtMenuName

End Sub

Private Sub txtMenuName_KeyPress(KeyAscii As Integer)
'
' Name:         txtMenuName_KeyPress
' Parameters:   KeyAscii            the key pressed
' Description:  Move focus to the next control when return is pressed.
'
   
    Dim Cancel As Boolean
    If KeyAscii = 13 Then
        Call txtMenuName_Validate(Cancel)
        If Not Cancel Then cboCategory.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtMenuName_Validate(Cancel As Boolean)
'
' Name:         txtMenuName_Validate
' Parameters:   Cancel          more a return value -- set TRUE to cancel validation
' Description:  Ensure the given name is legal before changing it.
'

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing Or Populating) Then
        With txtMenuName
            .Text = Trim(.Text)
            If Not (.Text = CurrentMenu.Name Or .Text = "") Then
                
                If (LCase(.Text) = LCase(CurrentItem.Tag)) Then
                    CurrentMenu.Name = .Text
                    CurrentItem.Tag = .Text
                    CurrentItem.Text = CurrentMenu.Name & ":"
                    MenuSet.DataChanged = True
                Else
                    Cancel = Not ValidateName(txtMenuName, MenuSet, CurrentMenu.Name)
                    If Not Cancel Then
                        MenuSet.SetMenuName CurrentMenu.Name, .Text
                        MenuSet.DataChanged = True
                        RefreshMenus CurrentMenu
                    End If
                End If
            
            Else
                .Text = CurrentMenu.Name
            End If
        End With
    End If
    
End Sub

Private Sub txtPath_GotFocus()
'
' Name:         txtPath_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtPath

End Sub

Private Sub txtSubmenuName_GotFocus()
'
' Name:         txtSubmenuName_GotFocus
' Description:  Select the text when this control gets focus.
'
    SelectText txtSubmenuName

End Sub

Private Sub txtSubmenuName_KeyPress(KeyAscii As Integer)
'
' Name:         txtSubmenuName_KeyPress
' Parameters:   KeyAscii            the key pressed
' Description:  Move focus to the next control when return is pressed.
'
   
    Dim Cancel As Boolean
    If KeyAscii = 13 Then
        Call txtSubmenuName_Validate(Cancel)
        If Not Cancel Then lstLinkedMenu.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtSubmenuName_Validate(Cancel As Boolean)
'
' Name:         txtSubmenuName_Validate
' Parameters:   Cancel          more a return value -- set TRUE to cancel validation
' Description:  Ensure the given name is legal before changing it.
'

    Dim CaseChange As Boolean

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing Or Populating) Then
        With txtSubmenuName
            
            .Text = Trim(.Text)
            If Not (.Text = CurrentItem.Tag Or .Text = "") Then
                
                Cancel = False
                CaseChange = (LCase(.Text) = LCase(CurrentItem.Tag))
                If Not CaseChange Then
                    Cancel = Not ValidateName(txtSubmenuName, CurrentMenu, CurrentItem.Tag)
                End If
                If Not Cancel Then
                    CurrentMenu.SetItemName CurrentItem.Tag, .Text
                    MenuSet.DataChanged = True
                    If CaseChange Then
                        CurrentItem.Tag = .Text
                        CurrentItem.Text = CurrentMenu.DisplayItem
                    Else
                        RefreshMenuItems CurrentMenu, .Text
                    End If
                End If
                
            Else
                .Text = CurrentItem.Tag
            End If
        
        End With
    End If
    
End Sub

Private Sub updCost_UpClick()
'
' Name:         updCost_UpClick
' Description:  Numerically increment a Cost field.  If it's not a number, it
'               becomes one.  Find the rightmost numeric value.
'

    Dim I As Integer
    Dim Cancel As Boolean
    
    If IsNumeric(txtItemCost.Text) Then
        txtItemCost.Text = CStr(CSng(txtItemCost.Text) + 1)
    Else
    
        For I = Len(txtItemCost.Text) To 1 Step -1
            If Not Mid(txtItemCost.Text, I, 1) Like "#" Then Exit For
        Next I
        txtItemCost.Text = Mid(txtItemCost.Text, I + 1)
        
    End If
    Call txtItemCost_Validate(Cancel)
    
End Sub

Private Sub updCost_DownClick()
'
' Name:         updCost_DownClick
' Description:  Numerically decrement a Cost field.  If it's not a number, it
'               becomes one.  Find the leftmost numeric value.
'

    Dim I As Integer
    Dim Cancel As Boolean
    
    If IsNumeric(txtItemCost.Text) Then
        txtItemCost.Text = CStr(CDbl(txtItemCost.Text) - 1)
    Else
    
        For I = 1 To Len(txtItemCost.Text)
            If Not Mid(txtItemCost.Text, I, 1) Like "#" Then Exit For
        Next I
        txtItemCost.Text = Left(txtItemCost.Text, I - 1)
        
    End If
    Call txtItemCost_Validate(Cancel)
    
End Sub

Private Sub updPosition_UpClick()
'
' Name:         updPosition_UpClick
' Description:  Move the selected node up in its order.
'

    Dim PrevItem As ListItem

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing) And InsideMenu Then
        If CurrentItem.Index > 2 Then
            Set PrevItem = lvwMenu.ListItems(CurrentItem.Index - 1)
            CurrentMenu.MoveTo CurrentItem.Tag
            If Not CurrentMenu.Off Then
                CurrentMenu.SwapFrontward
                MenuSet.DataChanged = True
                lvwMenu.ListItems.Remove CurrentItem.Index
                Set CurrentItem = Nothing
                Set CurrentItem = lvwMenu.ListItems.Add( _
                        Index:=PrevItem.Index, Text:=CurrentMenu.DisplayItem, _
                        Key:="k" & CurrentMenu.ItemName)
                CurrentItem.Tag = CurrentMenu.ItemName
                Set lvwMenu.SelectedItem = CurrentItem
                CurrentItem.EnsureVisible
            End If
        End If
    End If
    
End Sub

Private Sub updPosition_DownClick()
'
' Name:         updPosition_DownClick
' Description:  Move the selected node down in its order.
'

    Dim NextItem As ListItem

    If Not (CurrentMenu Is Nothing Or CurrentItem Is Nothing) And InsideMenu Then
        If CurrentItem.Index < lvwMenu.ListItems.Count Then
            Set NextItem = lvwMenu.ListItems(CurrentItem.Index + 1)
            CurrentMenu.MoveTo CurrentItem.Tag
            If Not CurrentMenu.Off Then
                CurrentMenu.SwapBackward
                MenuSet.DataChanged = True
                lvwMenu.ListItems.Remove CurrentItem.Index
                Set CurrentItem = Nothing
                Set CurrentItem = lvwMenu.ListItems.Add( _
                        Index:=NextItem.Index + 1, Text:=CurrentMenu.DisplayItem, _
                        Key:="k" & CurrentMenu.ItemName)
                CurrentItem.Tag = CurrentMenu.ItemName
                Set lvwMenu.SelectedItem = CurrentItem
                CurrentItem.EnsureVisible
            End If
        End If
    End If
    
End Sub

