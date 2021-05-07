VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Tag             =   "&Apply"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.CommandButton cmdLogBrowse 
         Caption         =   "Browse"
         Height          =   495
         Left            =   2415
         TabIndex        =   35
         Top             =   1695
         Width           =   1215
      End
      Begin VB.TextBox txtLogName 
         Height          =   345
         Left            =   225
         TabIndex        =   33
         Top             =   1170
         Width           =   5415
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "Keep Log File"
         Height          =   330
         Left            =   225
         TabIndex        =   32
         Top             =   390
         Width           =   1395
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "File name:"
         Height          =   195
         Left            =   225
         TabIndex        =   34
         Top             =   930
         Width           =   720
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.TextBox txtRetryIntvl 
         Height          =   330
         Left            =   660
         TabIndex        =   28
         Top             =   1935
         Width           =   945
      End
      Begin VB.TextBox txtRetryNum 
         Height          =   330
         Left            =   660
         TabIndex        =   27
         Top             =   2730
         Width           =   540
      End
      Begin VB.CheckBox chkRetry 
         Caption         =   "Retry bad IPs"
         Height          =   240
         Left            =   660
         TabIndex        =   26
         Top             =   1170
         Width           =   1530
      End
      Begin VB.Label Label9 
         Caption         =   $"frmOptions.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   810
         Left            =   270
         TabIndex        =   31
         Top             =   165
         Width           =   5265
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Retry interval (milliseconds 1000-60000)"
         Height          =   195
         Left            =   660
         TabIndex        =   30
         Top             =   1650
         Width           =   2805
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Number of retries (1 - 10)"
         Height          =   195
         Left            =   660
         TabIndex        =   29
         Top             =   2475
         Width           =   1740
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.ListBox lstEmail 
         Height          =   2400
         Left            =   1410
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   825
         Width           =   3165
      End
      Begin VB.TextBox txtEmail 
         Height          =   330
         Left            =   1410
         TabIndex        =   23
         Top             =   405
         Width           =   3165
      End
      Begin VB.CommandButton cmdRemoveEmail 
         Caption         =   "Remove"
         Height          =   495
         Left            =   270
         TabIndex        =   22
         Top             =   1200
         Width           =   945
      End
      Begin VB.CommandButton cmdAddEmail 
         Caption         =   "Add"
         Height          =   495
         Left            =   270
         TabIndex        =   21
         Top             =   510
         Width           =   945
      End
      Begin VB.CheckBox chkNotify 
         Caption         =   "Notify when IP goes bad"
         Height          =   225
         Left            =   1410
         TabIndex        =   20
         Top             =   3300
         Width           =   2070
      End
      Begin VB.CommandButton cmdTestEmail 
         Caption         =   "Send Test"
         Height          =   495
         Left            =   270
         TabIndex        =   19
         Top             =   1890
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Add Email address to notify list:"
         Height          =   195
         Left            =   1410
         TabIndex        =   25
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   225
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.TextBox txTTL 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1935
         Width           =   735
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "205.214.94.167"
         Top             =   495
         Width           =   3015
      End
      Begin VB.TextBox txtIntvl 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.ListBox lstPing 
         Height          =   2985
         Left            =   3435
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   495
         Width           =   2040
      End
      Begin VB.CommandButton cmdAddIp 
         Caption         =   "Add"
         Height          =   495
         Left            =   135
         TabIndex        =   9
         Top             =   3015
         Width           =   945
      End
      Begin VB.CommandButton cmdRemoveIp 
         Caption         =   "Remove"
         Height          =   495
         Left            =   1260
         TabIndex        =   8
         Top             =   3015
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time To Live (TTL) :"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Add IP Address to ping list:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Ping Interval (seconds):"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   975
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ping List:"
         Height          =   195
         Left            =   3420
         TabIndex        =   15
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "1...255"
         Height          =   195
         Left            =   915
         TabIndex        =   14
         Top             =   1980
         Width           =   495
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "IP Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Email List"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Retry"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Log File"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
  Call SetOptions
  Call SetTab
End Sub

Private Sub SetOptions()
  Call LogOptions
  Call IPOptions
  Call EmailOptions
  Call RetryOptions
  Call UserSettings.SaveSettings
End Sub

Private Sub cmdCancel_Click()
  frmMain.SetFocus
  Unload Me
End Sub

Private Sub cmdLogBrowse_Click()
  Dim FileDlg As OPENFILENAME
  Dim lngReturn As Long
  
  With FileDlg
    .lStructSize = Len(FileDlg)
    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrFilter = "All Files"
    .FileName = Space(254)
    .nMaxFile = 255
    .lpstrFileTitle = Space(254)
    .nMaxFileTitle = 255
    .lpstrInitialDir = App.Path
    .lpstrTitle = "Locate Log File"
    .flags = cdlOFNHideReadOnly
  
    lngReturn = GetOpenFileName(FileDlg)
    
    If lngReturn >= 1 Then
       txtLogName.Text = .FileName
    End If
  End With
  
End Sub

Private Sub cmdOK_Click()
  Call SetOptions
  frmMain.SetFocus
  Unload Me
End Sub

Private Sub Form_Activate()
  Call SetTxtBox(txtIP)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer
  i = tbsOptions.SelectedItem.index
  'handle ctrl+tab to move to the next tab
  If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
    If i = tbsOptions.Tabs.Count Then
      'last tab so we need to wrap to tab 1
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
    Else
      'increment the tab
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
    End If
  ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
    If i = 1 Then
      'last tab so we need to wrap to tab 1
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
    Else
      'increment the tab
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
    End If
  End If
End Sub

Private Sub Form_Load()
  Call LoadEmail
  Call LoadIPSettings
  Call LoadRetry
  Call LoadLog
End Sub

Private Sub tbsOptions_Click()
  Dim i As Integer
  'show and enable the selected tab's controls
  'and hide and disable all others
  For i = 0 To tbsOptions.Tabs.Count - 1
    If i = tbsOptions.SelectedItem.index - 1 Then
      picOptions(i).Left = 210
      picOptions(i).Enabled = True
    Else
      picOptions(i).Left = -20000
      picOptions(i).Enabled = False
    End If
  Next
  
  Call SetTab
End Sub

Private Sub SetTab()
  With tbsOptions
    Select Case .SelectedItem
      Case .Tabs(1).Caption
        Call SetTxtBox(txtIP)
  
      Case .Tabs(2).Caption
        Call SetTxtBox(txtEmail)
    End Select
  End With
End Sub
Private Sub SetTxtBox(tb As TextBox)
  On Error Resume Next
  With tb
    .SetFocus
    .SelStart = Len(tb)
  End With
End Sub
Private Sub RetryOptions()
  If txtRetryIntvl < 1000 Then
    txtRetryIntvl = 1000
  ElseIf txtRetryIntvl > 60000 Then
    txtRetryIntvl = 60000
  End If
  
  'frmMain.tmRetry.Interval = txtRetryIntvl
  
  If txtRetryNum < 1 Then
    txtRetryNum = 1
  ElseIf txtRetryNum > 10 Then
    txtRetryNum = 10
  End If
  
  With UserSettings
    .Retry = (chkRetry.Value = vbChecked)
    .RetryIntvl = txtRetryIntvl
    .RetryNum = txtRetryNum
  End With
End Sub
Private Sub EmailOptions()
  UserSettings.Notify = (chkNotify.Value = vbChecked)
End Sub

Private Sub IPOptions()
  UserSettings.Interval = txtIntvl * 1000
  With frmMain
    If .cmdPause.Caption = "PAUSE" Then 'not paused
      .Timer1.Enabled = False
      Call .SetupTimer(txtIntvl * 1000)
      .Timer1.Enabled = True
      Call .ResetNextCounter
    End If
  End With
  If txTTL < 1 Then
    txTTL = 1
  ElseIf txTTL > 255 Then
    txTTL = 255
  End If
  
  UserSettings.TTL = txTTL
End Sub
Private Sub LogOptions()
  With UserSettings
    .KeepLog = (chkLog.Value = vbChecked)
    .LogName = txtLogName
  End With
End Sub
Private Sub cmdAddIp_Click()
  Dim intX As Integer
  If Len(txtIP) < 1 Then GoTo Exit_Here
  
  With lstPing
    For intX = 0 To .ListCount - 1
      If .List(intX) = txtIP Then GoTo Exit_Here
    Next
    .AddItem txtIP
    
    Call UserSettings.InitList(.ListCount)
    
    For intX = 0 To .ListCount - 1
      If .List(intX) = txtIP Then .Selected(intX) = True
      UserSettings.Checked(intX) = .Selected(intX)
      UserSettings.List(intX) = .List(intX)
    Next
  End With

Exit_Here:
  Call SetTxtBox(txtIP)
End Sub

Private Sub cmdRemoveIp_Click()
  Dim intX As Integer
  
  With lstPing
    If .ListIndex = -1 Then GoTo Exit_Here
    .RemoveItem .ListIndex
    
    Call UserSettings.InitList(.ListCount)
    
    For intX = 0 To .ListCount - 1
      UserSettings.Checked(intX) = .Selected(intX)
      UserSettings.List(intX) = .List(intX)
    Next
  End With

Exit_Here:
  Call SetTxtBox(txtIP)
End Sub

Private Sub cmdRemoveEmail_Click()
  Dim intX As Integer
  
  With lstEmail
    If .ListIndex = -1 Then GoTo Exit_Here
    .RemoveItem .ListIndex
    
    Call UserSettings.InitEmail(.ListCount)
    
    For intX = 0 To .ListCount - 1
      UserSettings.Email(intX) = .List(intX)
    Next
  End With

Exit_Here:
  Call SetTxtBox(txtEmail)
End Sub

Private Sub cmdAddEmail_Click()
  Dim intX As Integer
  If Len(txtEmail) < 1 Then GoTo Exit_Here
  
  With lstEmail
    For intX = 0 To .ListCount - 1
      If .List(intX) = txtEmail Then GoTo Exit_Here
    Next
    .AddItem txtEmail
    
    Call UserSettings.InitEmail(.ListCount)
    
    For intX = 0 To .ListCount - 1
      UserSettings.Email(intX) = .List(intX)
    Next
  End With
  txtEmail = vbNullString

Exit_Here:
  Call SetTxtBox(txtEmail)
End Sub

Private Sub cmdTestEmail_Click()
  Dim msg As String
  Dim intX As Integer
  
  Screen.MousePointer = vbHourglass
  
  msg = "This is a test email from the Ping Monitor utility."
  
  With frmMain
    .mapiLogOn.SignOn ' use current user
  
    Do While .mapiLogOn.SessionID = 0
      DoEvents ' need to wait until the new session is created
    Loop
    
    For intX = 0 To UserSettings.EmailCount - 1
      Call .SendToEmail(UserSettings.Email(intX), msg)
    Next
    
    .mapiLogOn.SignOff
  End With
  
  Screen.MousePointer = vbNormal
  
  Call SetTxtBox(txtEmail)

End Sub

Private Sub LoadEmail()
  Dim intX As Integer
  
  With UserSettings
    chkNotify.Value = .Notify * -1
    For intX = 0 To .EmailCount - 1
      lstEmail.AddItem .Email(intX)
    Next
  End With
  
  With lstEmail
    If .ListIndex > -1 Then txtEmail = .List(.ListIndex)
  End With
End Sub

Private Sub LoadIPSettings()
  Dim intX As Integer
  
  With UserSettings
    chkLog.Value = .KeepLog * -1
    txTTL = .TTL
    txtIntvl = .Interval \ 1000
    For intX = 0 To .ListCount - 1
      lstPing.AddItem .List(intX)
      lstPing.Selected(intX) = .Checked(intX)
    Next
  End With
  
  With lstPing
    If .ListIndex > -1 Then txtIP = .List(.ListIndex)
  End With
End Sub
Private Sub LoadLog()
  With UserSettings
    chkLog.Value = .KeepLog * -1
    txtLogName = .LogName
  End With
End Sub
Private Sub LoadRetry()
  With UserSettings
    chkRetry.Value = .Retry * -1
    txtRetryIntvl = .RetryIntvl
    txtRetryNum = .RetryNum
  End With
End Sub
Private Sub lstPing_ItemCheck(Item As Integer)
  UserSettings.Checked(Item) = lstPing.Selected(Item)
End Sub

Private Sub lstPing_Click()
  With lstPing
    txtIP = .List(.ListIndex)
  End With
  Call SetTxtBox(txtIP)
End Sub

Private Sub lstEmail_Click()
  With lstEmail
    txtEmail = .List(.ListIndex)
  End With
  Call SetTxtBox(txtEmail)
End Sub
