VERSION 5.00
Begin VB.Form frmSaveDirectory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose Folder"
   ClientHeight    =   3600
   ClientLeft      =   8460
   ClientTop       =   6090
   ClientWidth     =   5745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save in this Folder"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chkPrompt 
      Caption         =   "&Prompt for Confirmation Before Overwriting Files"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.DirListBox dirDirBox 
      Height          =   1665
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.DriveListBox drvDriveBox 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdSaveIndividually 
      Caption         =   "Save Files Individually"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblCurDir 
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
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   5295
   End
   Begin VB.Label lblCaption 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblHeader 
      Caption         =   "&Choose a Folder in which to Save"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmSaveDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum SaveDirectoryType
    sdOK
    sdCancel
    sdIndividual
End Enum

Public Value As SaveDirectoryType   'Whether the user clicks OK, Cancel, or Save Individually
Public Overwrite As Boolean         'Whether the user choses to automatically overwrite files

Public Sub GetSaveDirectory(What As String)
'
' Name:         GetSaveDirectory
' Parameters:   What        the purpose for which this window is used
' Description:  Initialize and show the window with an appropriate caption.
'

    chkPrompt = GetSetting(App.Title, "Output", "OverwritePrompt", vbUnchecked)
    lblCaption = What
    Me.Show vbModal

End Sub

Private Sub chkPrompt_Click()
'
' Name:         chkPrompt_Click
' Description:  Remember that the user wants to be prompted for confirmation before
'               overwriting files
'

    SaveSetting App.Title, "Output", "OverwritePrompt", chkPrompt
    Overwrite = Not (chkPrompt = vbChecked)

End Sub

Private Sub cmdCancel_Click()
'
' Name:         cmdCancel_Click
' Description:  Cancel and dismiss the window.
'

    Value = sdCancel
    Me.Hide

End Sub

Private Sub cmdOK_Click()
'
' Name:         cmdOK_Click
' Description:  Save the chosen directory and dismiss the window.
'

    Value = sdOK
    SaveSetting App.Title, "Files", "ExportDir", CurDir
    Me.Hide
    
End Sub

Private Sub cmdSaveIndividually_Click()
'
' Name:         cmdSaveIndividually_Click
' Description:  The user chooses to save files individually; dismiss the window.
'

    Value = sdIndividual
    Me.Hide

End Sub

Private Sub dirDirBox_Change()
'
' Name:         dirDirBox_Change
' Description:  Change to the selected directory.
'

    On Error GoTo dirDirBoxChangeError
    
    lblCurDir = dirDirBox.Path
    ChDir dirDirBox.Path
    
    GoTo dirDirBoxEnd
    
dirDirBoxChangeError:
    If MsgBox(Err.Description, vbRetryCancel + vbExclamation, "Invalid Path") _
            = vbCancel Then
        Resume dirDirBoxEnd
    Else
        Resume
    End If
    
dirDirBoxEnd:
    
End Sub

Private Sub drvDriveBox_Change()
'
' Name:         drvDriveBox_Change
' Description:  Change to the selected drive.
'

    On Error GoTo drvDriveBoxChangeError
    
    ChDrive drvDriveBox.Drive
    dirDirBox.Path = CurDir
    
    GoTo drvDriveBoxEnd
    
drvDriveBoxChangeError:
    If MsgBox(Err.Description, vbRetryCancel + vbExclamation, "Invalid Path") _
            = vbCancel Then
        drvDriveBox.Drive = CurDir
        Resume drvDriveBoxEnd
    Else
        Resume
    End If
    
drvDriveBoxEnd:
    
End Sub

Private Sub Form_Load()
'
' Name:         Form_Load
' Description:  Initialize the window with the user's previous preferences.
'

    On Error GoTo FormLoadError
    dirDirBox.Path = GetSetting(App.Title, "Files", "ExportDir", CurDir)
    drvDriveBox.Drive = dirDirBox.Path
    lblCurDir = dirDirBox.Path
    ChDir dirDirBox.Path
        
    GoTo FormLoadEnd
    
FormLoadError:
    If MsgBox(Err.Description, vbRetryCancel + vbExclamation, "Invalid Path") _
            = vbCancel Then
        drvDriveBox.Drive = "C:\"
        dirDirBox.Path = "C:\"
        lblCurDir = "C:\"
        Resume FormLoadEnd
    Else
        Resume
    End If
    
FormLoadEnd:

End Sub

