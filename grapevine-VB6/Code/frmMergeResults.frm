VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMergeResults 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Merge Results"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "entity"
         Text            =   "Entity"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "action"
         Text            =   "Action"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblCount 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
End
Attribute VB_Name = "frmMergeResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowResults(ByVal Results As String, ParentForm As Form)
'
' Name:         ShowResults
' Parameters:   Results         vbCr-delimited string of results
'               Parent Form     parent form for show modal
' Description:  Modally show the results of a merge operation or exchange load.
'

    Dim NewItem As ListItem
    Dim Delim As Integer
    
    If Results = "" Then
        
        lvwResults.ListItems.Add , , "No Changes"
        
    Else
    
        Delim = InStr(Results, vbCr)
        Do Until Delim = 0
        
            Set NewItem = lvwResults.ListItems.Add(, , Left(Results, Delim - 1))
            Results = Mid(Results, Delim + 1)
            Delim = InStr(Results, vbCr)
            If Delim > 0 Then
                NewItem.ListSubItems.Add , , Left(Results, Delim - 1)
                Results = Mid(Results, Delim + 1)
                Delim = InStr(Results, vbCr)
            End If
        
        Loop
    
        lblCount.Caption = CStr(lvwResults.ListItems.Count) & " entities loaded."
    
    End If
    
    Me.Show vbModal, ParentForm

End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Unload this form.
'

    Unload Me

End Sub
