VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "View Log"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstLog 
      Height          =   450
      Left            =   1155
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "Delete"
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  Call LoadLogFile
End Sub

Private Sub Form_Resize()
  lstLog.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub LoadLogFile()
  Dim handle As Integer
  handle = FreeFile
  On Error GoTo NoFile
  
  Screen.MousePointer = vbHourglass
  With lstLog
    Dim strLine As String
    .Visible = False
    .Clear
    Open UserSettings.LogName For Input As #handle
      Do Until EOF(1)
        Line Input #1, strLine
        .AddItem strLine
      Loop
    Close
    .Visible = True
    .Selected(.ListCount - 1) = True
  End With

Exit_Here:
  Screen.MousePointer = vbDefault
  Exit Sub

NoFile:
  frmMain.cmdLog.Enabled = False
  Resume Exit_Here
End Sub

Private Sub mnuDelete_Click()
  Dim lngC As Long
  lngC = MsgBox("Are you Sure?", vbOKCancel, "Delete Log File?")
  If lngC = vbOK Then Kill UserSettings.LogName
End Sub
