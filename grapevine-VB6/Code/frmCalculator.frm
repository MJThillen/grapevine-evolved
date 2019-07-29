VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point Counting Aid"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmCalculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   7695
   Visible         =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1095
   End
   Begin MSComCtl2.UpDown updPoints 
      Height          =   285
      Left            =   6930
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2640
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtPoints"
      BuddyDispid     =   196611
      OrigLeft        =   5400
      OrigTop         =   2640
      OrigRight       =   5895
      OrigBottom      =   2925
      Max             =   9999
      Min             =   -9999
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.ListBox lstTraits 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   2010
      Left            =   4680
      TabIndex        =   11
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtPoints 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Text            =   "0"
      Top             =   2640
      Width           =   570
   End
   Begin MSComctlLib.ListView lvwInvoice 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Content"
         Object.Width           =   4357
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1032
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Change"
         Object.Width           =   1376
      EndProperty
   End
   Begin VB.Image imgRace 
      Height          =   255
      Index           =   1
      Left            =   4200
      Top             =   240
      Width           =   255
   End
   Begin VB.Image imgRace 
      Height          =   255
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblLabel 
      Caption         =   $"frmCalculator.frx":058A
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   6840
      X2              =   6360
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Image imgScales 
      Height          =   480
      Index           =   2
      Left            =   4680
      Picture         =   "frmCalculator.frx":063B
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgScales 
      Height          =   480
      Index           =   1
      Left            =   5220
      Picture         =   "frmCalculator.frx":0F05
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgScales 
      Height          =   480
      Index           =   0
      Left            =   5760
      Picture         =   "frmCalculator.frx":17CF
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBalance 
      Height          =   480
      Left            =   6960
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblXPTotal 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "XP Spent ="
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblCharName 
      Alignment       =   2  'Center
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblListCaption 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Points Spent ="
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblPointTotal 
      Alignment       =   2  'Center
      Caption         =   "-"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblListPoints 
      Caption         =   "&Points spent above ="
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      Top             =   2670
      Width           =   1695
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IMG_LEFT = 0
Private Const IMG_BALANCE = 1
Private Const IMG_RIGHT = 2
Private Const PKEY = "points"
Private Const CKEY = "change"

Private Character As Object
Private Populating As Boolean

Public Sub ShowCalculator(Char As Object)
'
' Name:         ShowCalculator
' Parameters:   Character           Character object who is to be calculated
' Description:  Estimate and calculate the points invested in a character.
'

    Dim PointList As LinkedTraitList
    Dim NewItem As ListItem
    Dim KeyTitle As String
    Dim Value As Variant
    Dim Count As Single
    Dim Spending As Single
    Dim Total As Double
    Dim P As Integer
    Dim S As Integer
    Dim M As Integer
    Dim NegList As Boolean
    
    Set Character = Char
    lblCharName.Caption = Character.Name
    imgRace(0).Picture = mdiMain.imlSmallIcons.ListImages(Character.Race).Picture
    imgRace(1).Picture = imgRace(0).Picture
    Set PointList = New LinkedTraitList
    lvwInvoice.ListItems.Clear
    
    P = Character.PhysicalList.Count
    S = Character.SocialList.Count
    M = Character.MentalList.Count
    
    If P <= S And S <= M Then
        P = 3: S = 5: M = 7
    ElseIf P <= M And M <= S Then
        P = 3: M = 5: S = 7
    ElseIf S <= P And P <= M Then
        S = 3: P = 5: M = 7
    ElseIf S <= M And M <= P Then
        S = 3: M = 5: P = 7
    ElseIf M <= P And P <= S Then
        M = 3: P = 5: S = 7
    Else
        M = 3: S = 5: P = 7
    End If
    
    PointList.SetAlphabetized False
    PointList.Append qkPhysical, , CStr(P)
    PointList.Append qkPhysicalNeg
    PointList.Append qkSocial, , CStr(S)
    PointList.Append qkSocialNeg
    PointList.Append qkMental, , CStr(M)
    PointList.Append qkMentalNeg
    PointList.Append qkAbilities, , "5"
    PointList.Append qkBackgrounds, , "5"
    PointList.Append qkInfluences

    On Error Resume Next
    Character.CompleteEstimateList PointList
    On Error GoTo 0

    PointList.First
    Do Until PointList.Off
        
        Character.GetValue PointList.Trait.Name, Value
        If IsObject(Value) Then
            Count = Value.Count
            Spending = Value.TraitCount
            NegList = Value.Negative
        Else
            Count = CSng(Val(Value))
            Spending = Count
            NegList = (PointList.Trait.Number < 0)
        End If
        Spending = CSng((Spending - Val(PointList.Trait.Note)) * PointList.Trait.Number)
        
        KeyTitle = CStr(Game.QueryEngine.KeysToTitles(PointList.Trait.Name))
        If InStr(KeyTitle, " (") > 0 Then KeyTitle = Left(KeyTitle, InStr(KeyTitle, " (") - 1)
        KeyTitle = CStr(Count) & " " & KeyTitle
        
        Set NewItem = lvwInvoice.ListItems.Add(Text:=KeyTitle)
        NewItem.Tag = PointList.Trait.Name
        NewItem.ListSubItems.Add Key:=PKEY, Text:=CStr(Spending)
        NewItem.ListSubItems.Add Key:=CKEY, Text:=IIf(NegList, "earned", "spent")
        
        If NegList Then Spending = -Spending
        Total = Total + Spending
        
        PointList.MoveNext
    Loop
    
    Set NewItem = lvwInvoice.ListItems.Add(Text:="Free Points")
    NewItem.ListSubItems.Add Key:=PKEY, Text:="5"
    NewItem.ListSubItems.Add Key:=CKEY, Text:="earned"
    Total = Total - 5
    
    Set NewItem = lvwInvoice.ListItems.Add(Text:="Additional Costs")
    NewItem.ListSubItems.Add Key:=PKEY, Text:="0"
    NewItem.ListSubItems.Add Key:=CKEY, Text:="spent"
    
    Set PointList = Nothing
    
    lblXPTotal.Caption = Character.Experience.Earned - Character.Experience.Unspent
    lblPointTotal.Caption = Total
    Set lvwInvoice.SelectedItem = lvwInvoice.ListItems(1)
    Call lvwInvoice_ItemClick(lvwInvoice.SelectedItem)
    
    Me.Show
    Me.SetFocus
    
End Sub

Public Sub SetDefaultOutput()
'
' Name:         SetDefaultOutput
' Description:  Prepare an output setup for frmOutput.
'
    With OutputEngine
        .Template = tnXPHistory
        .SelectSet(osCharacters).Clear
        .SelectSet(osCharacters).Add Character.Name
        .GameDate = 0
    End With

End Sub

Private Sub cmdClose_Click()
'
' Name:         cmdClose_Click
' Description:  Unload this window
'

    Unload Me
    
End Sub

Private Sub Form_Activate()
'
' Name:         From_Activate
' Description:  Refresh the trait list.
'
    
    CharacterList.MoveTo Character.Name
    If CharacterList.Off Then
        Set Character = Nothing
        Unload Me
    Else
        Call lvwInvoice_ItemClick(lvwInvoice.SelectedItem)
    End If
    
End Sub

Private Sub lblPointTotal_Change()
'
' Name:         lblPointTotal_Change
' Description:  Display the needed picture.
'

    Select Case (Val(lblPointTotal) - Val(lblXPTotal))
        Case Is < 0
            imgBalance.Picture = imgScales(IMG_LEFT).Picture
        Case 0
            imgBalance.Picture = imgScales(IMG_BALANCE).Picture
        Case Is > 0
            imgBalance.Picture = imgScales(IMG_RIGHT).Picture
    End Select
    
End Sub

Private Sub lvwInvoice_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
' Name:         lvwInvoice_ItemClick
' Description:  Update the list at right describing the selected item.
'

    lstTraits.Clear

    If Not Item Is Nothing Then
        
        lblListCaption.Caption = Item.Text
        
        If Not Item.Tag = "" Then

            Dim Value As Variant
            Dim KeyTitle As String
            
            Character.GetValue Item.Tag, Value

            KeyTitle = CStr(Game.QueryEngine.KeysToTitles(Item.Tag))
            If InStr(KeyTitle, " (") > 0 Then KeyTitle = Left(KeyTitle, InStr(KeyTitle, " (") - 1)

            If IsObject(Value) Then
                Dim List As LinkedTraitList
                Set List = Value
                List.First
                Do Until List.Off
                    lstTraits.AddItem List.DisplayTrait
                    List.MoveNext
                Loop
                KeyTitle = CStr(List.Count) & " " & KeyTitle
            Else
                lstTraits.AddItem String(Int(Val(Value)), "O")
            End If
            
            If Item.Text <> KeyTitle Then Item.Text = KeyTitle
            lblListCaption.Caption = KeyTitle
            
        End If
        
        Populating = True
        lblListPoints.Caption = IIf(Item.ListSubItems(CKEY).Text = "earned", _
                "&Points earned above =", "&Points spent above =")
        txtPoints.Text = Item.ListSubItems(PKEY).Text
        Populating = False
        
    End If

End Sub

Private Sub txtPoints_Change()
'
' Name:         txtPoints_Change
' Description:  Adjust the points assumed to have been spent.
'

    If Not (Populating Or lvwInvoice.SelectedItem Is Nothing) Then
    
        Dim OldTotal As Single
        Dim OldValue As Single
        Dim NewValue As Single
        
        NewValue = CSng(Val(txtPoints.Text))
        OldValue = CSng(Val(lvwInvoice.SelectedItem.ListSubItems(PKEY).Text))
        OldTotal = CSng(Val(lblPointTotal.Caption))
    
        lvwInvoice.SelectedItem.ListSubItems(PKEY).Text = NewValue
        If lvwInvoice.SelectedItem.ListSubItems(CKEY).Text = "earned" Then
            OldValue = -OldValue
            NewValue = -NewValue
        End If
        lblPointTotal.Caption = CStr(OldTotal - OldValue + NewValue)
    
    End If

End Sub

Private Sub txtPoints_GotFocus()
'
' Name:         txtPoints_GotFocus
' Description:  Select the text.
'
    SelectText txtPoints
    
End Sub

Private Sub txtPoints_KeyPress(KeyAscii As Integer)
'
' Name:         txtpoints_KeyPress
' Description:  Move to the next field when enter is pressed.
'

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not lvwInvoice.SelectedItem Is Nothing Then
            If lvwInvoice.SelectedItem.Index < lvwInvoice.ListItems.Count Then
                Set lvwInvoice.SelectedItem = lvwInvoice.ListItems(lvwInvoice.SelectedItem.Index + 1)
                Call lvwInvoice_ItemClick(lvwInvoice.SelectedItem)
            End If
        End If
    End If

End Sub
