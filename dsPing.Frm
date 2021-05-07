VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Ping Monitor"
   ClientHeight    =   5010
   ClientLeft      =   1755
   ClientTop       =   2325
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dsPing.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5010
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLog 
      Caption         =   "View Log"
      Height          =   495
      Left            =   2573
      TabIndex        =   11
      ToolTipText     =   "Force a ping"
      Top             =   2985
      Width           =   690
   End
   Begin VB.Timer tmRetry 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3660
      Top             =   2625
   End
   Begin VB.Timer tmNext 
      Interval        =   1000
      Left            =   3630
      Top             =   2970
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   240
      Top             =   4290
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession mapiLogOn 
      Left            =   1065
      Top             =   4305
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.PictureBox picReturn 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   3870
      TabIndex        =   8
      Top             =   3975
      Width           =   3870
      Begin VB.ListBox lbReturn 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ping Return Info"
         Top             =   0
         Width           =   3030
      End
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "PING"
      Height          =   495
      Left            =   1590
      TabIndex        =   5
      ToolTipText     =   "Force a ping"
      Top             =   2985
      Width           =   690
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "PAUSE"
      Height          =   495
      Left            =   608
      TabIndex        =   4
      ToolTipText     =   "Pause/resume ping timer"
      Top             =   2985
      Width           =   690
   End
   Begin VB.ListBox lstBadIP 
      Height          =   2580
      Left            =   2010
      TabIndex        =   3
      ToolTipText     =   "Bad Pings"
      Top             =   300
      Width           =   1725
   End
   Begin VB.ListBox lstGoodIP 
      Height          =   2580
      Left            =   105
      TabIndex        =   2
      ToolTipText     =   "Good Pings"
      Top             =   300
      Width           =   1725
   End
   Begin VB.Timer Timer1 
      Left            =   3645
      Top             =   3405
   End
   Begin VB.Label lblNext 
      AutoSize        =   -1  'True
      Caption         =   "next"
      Height          =   210
      Left            =   3075
      TabIndex        =   10
      Top             =   3645
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "FAILURE:"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2528
      TabIndex        =   7
      Top             =   90
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "SUCCESS:"
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   593
      TabIndex        =   6
      Top             =   90
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Ping:"
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   3645
      Width           =   705
   End
   Begin VB.Label lblNow 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   855
      TabIndex        =   0
      Top             =   3645
      Width           =   45
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PingState
  [Completed] = 0
  [Failed to Complete] = 1
  [Disconnected] = 2
End Enum

'******************************************************************************
Private Const IP_STATUS_BASE = 11000

Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
'
'   The next group are status codes passed up on status indications to
'   transport layer protocols.
'
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)

Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)

'   option information for network ping, we don't implement these here as this is
'   a simple sample (simon says).
Private Type ip_option_information
    TTL             As Byte      'Time To Live
    Tos             As Byte     'Type Of Service
    flags           As Byte     'IP header flags
    OptionsSize     As Byte     'Size in bytes of options data
    OptionsData     As Long     'Pointer to options data
End Type

'   structure that is returned from the ping to give status and error information
Private Type icmp_echo_reply
    Address         As Long             'Replying address
    Status          As Long             'Reply IP_STATUS, values as defined above
    RoundTripTime   As Long             'RTT in milliseconds
    DataSize        As Integer          'Reply data size in bytes
    Reserved        As Integer          'Reserved for system use
    DataPointer     As Long             'Pointer to the reply data
    Options         As ip_option_information    'Reply options
    Data            As String * 250     'Reply data which should be a copy of the string sent, NULL terminated
                                        ' this field length should be large enough to contain the string sent
End Type

'   declares for function to be used from icmp.dll
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

Private Declare Function IcmpSendEcho Lib "icmp.dll" _
(ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, _
ByVal RequestData As String, ByVal RequestSize As Integer, _
RequestOptions As ip_option_information, ReplyBuffer As icmp_echo_reply, _
ByVal ReplySize As Long, ByVal Timeout As Long) As Long

Private Const PING_TIMEOUT = 200        ' number of milliseconds to wait for the reply

Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYSSTATUS_LEN = 256
Private Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Private Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Private Const SOCKET_ERROR = -1

Private Type tagWSAData
        wVersion            As Integer
        wHighVersion        As Integer
        szDescription       As String * WSADESCRIPTION_LEN_1
        szSystemStatus      As String * WSASYSSTATUS_LEN_1
        iMaxSockets         As Integer
        iMaxUdpDg           As Integer
        lpVendorInfo        As String * 200
        End Type
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As tagWSAData) As Integer
Private Declare Function WSACleanup Lib "wsock32" () As Integer

Private Sub btnExit_Click()
   Unload Me
   End
End Sub

Private Function PingIp(ByVal IP As String) As PingState

    Dim hFile       As Long             ' handle for the icmp port opened
    Dim lRet        As Long             ' hold return values as required
    Dim lIPAddress  As Long
    Dim strMessage  As String
    Dim pOptions    As ip_option_information
    Dim pReturn     As icmp_echo_reply
    Dim iVal        As Integer
    Dim lPingRet    As Long
    Dim pWsaData    As tagWSAData
    
    strMessage = "Echo this string of data"
    
    iVal = WSAStartup(&H101, pWsaData)
    
    '   convert the IP address to a long, lIPAddress will be zero
    '   if the function failed. Normally you wouldn't ping if the address
    '   was no good to start with but we don't mind seeing bad return status
    '   as that is what samples are all about
    lIPAddress = ConvertIPAddressToLong(IP)
    
    '   open up a file handle for doing the ping
    hFile = IcmpCreateFile()
    
    '   set the TTL from the text box, try values of 1 to 255
    pOptions.TTL = UserSettings.TTL
    
    '   Call the function that actually does the ping. It is a blocking call so we
    '   don't get control back until it completes.
    lRet = IcmpSendEcho(hFile, _
                        lIPAddress, _
                        strMessage, _
                        Len(strMessage), _
                        pOptions, _
                        pReturn, _
                        Len(pReturn), _
                        PING_TIMEOUT)

  If lRet = 0 Then
        ' the ping failed for some reason, hopefully the error is in the return buffer
        lbReturn.AddItem "[" & IP & "] failed with error " & pReturn.Status
        lbReturn.ListIndex = lbReturn.ListCount - 1
        PingIp = Disconnected
    Else
        ' the ping succeeded, .Status will be 0, .RoundTripTime is the time in ms for
        '   the ping to complete, .Data is the data returned (NULL terminated), .Address
        '   is the Ip address that actually replied, .DataSize is the size of the string in
        '   .Data
        If pReturn.Status <> 0 Then
            lbReturn.AddItem "Error -> [" & IP & "] failed to complete, code = " & pReturn.Status
            lbReturn.ListIndex = lbReturn.ListCount - 1
            PingIp = [Failed to Complete]
        Else
            lbReturn.AddItem "[" & IP & "] Success -> completion time is " & pReturn.RoundTripTime & "ms."
            lbReturn.ListIndex = lbReturn.ListCount - 1
            PingIp = Completed
        End If
    End If
                        
    '   close the file handle that was used
    lRet = IcmpCloseHandle(hFile)
    
    iVal = WSACleanup()
    

End Function
'
'   ConvertIPAddressToLong
'
'   Converts a dotted IP address (eg: "123.234.2.45") to a long
'   integer for use in sending a ping. This routine converts
'   the string as required by an Intel system.
'
'   Essentially we take the 4 numbers, flip them around and make
'   a long by shifting all the parts into the correct byte. We
'   do it here by making a hex string and converting it to a long.
'   Not pretty but it works (most of the time<g>).
'
'   When we get in "a.b.c.d" what we want out is Val(&Hddccbbaa).
'

Function ConvertIPAddressToLong(strAddress As String) As Long

    Dim strTemp             As String
    Dim lAddress            As Long
    Dim iValCount           As Integer
    Dim lDotValues(1 To 4)  As String
    
    ' set up the initial storage and counter
    strTemp = strAddress
    iValCount = 0
    
    ' keep going while we still have dots in the string
    While InStr(strTemp, ".") > 0
        iValCount = iValCount + 1   ' count the number
        lDotValues(iValCount) = Mid(strTemp, 1, InStr(strTemp, ".") - 1)    ' pick it off and convert it
        strTemp = Mid(strTemp, InStr(strTemp, ".") + 1) ' chop off the number and the dot
        Wend
        
    ' the string only has the last number in it now
    iValCount = iValCount + 1
    lDotValues(iValCount) = strTemp
    
    ' if we didn't get four pieces then the IP address is no good
    If iValCount <> 4 Then
        ConvertIPAddressToLong = 0
        Exit Function
        End If
        
    '   take the four value, hex them, pad to 2 digits, make a hex
    '   string and then convert the whole mess to a long for returning
    lAddress = Val("&H" & Right("00" & Hex(lDotValues(4)), 2) & _
                Right("00" & Hex(lDotValues(3)), 2) & _
                Right("00" & Hex(lDotValues(2)), 2) & _
                Right("00" & Hex(lDotValues(1)), 2))
                
    '   set the return value
    ConvertIPAddressToLong = lAddress
    
End Function

Private Sub cmdLog_Click()
  Call frmLog.Show
End Sub

Private Sub cmdPause_Click()
  With cmdPause
    If .Caption = "PAUSE" Then
      .Caption = "RESUME"
      Timer1.Interval = 0
      tmNext.Interval = 0
    Else
      .Caption = "PAUSE"
      tmNext.Interval = 1000
      Call ResetNextCounter
      Call SetupTimer(UserSettings.Interval)
    End If
  End With
End Sub


Private Sub Form_Load()

  Caption = Caption & GetAppVer
  
  With UserSettings
    Call .GetSettings
    Call SetupTimer(.Interval)

    Call SendPing
    cmdLog.Enabled = (Len(Dir(.LogName)) > 1)
  End With
  
  Call ResetNextCounter

End Sub
Public Sub SetupTimer(ByVal Intvl As Long)
  If Intvl < 60000 Then
    Timer1.Interval = Intvl
  Else
    Timer1.Interval = 60000
  End If
End Sub
Private Sub Form_Resize()

  picReturn.Visible = False
  With lblNow
    picReturn.Top = .Top + .Height + 50
  End With
  
  On Error Resume Next
  With picReturn
    .Height = Me.ScaleHeight - .Top
    lbReturn.Move 0, 0, .Width, .Height
  End With
  picReturn.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload frmOptions
  Unload frmLog
  Set UserSettings = Nothing
End Sub

Private Sub mnuOptions_Click()
  frmOptions.Show
End Sub

Private Sub Timer1_Timer()
  Static iElapsedMin As Integer
  Dim vMins As Integer
  
  vMins = (UserSettings.Interval \ 1000) \ 60

  If vMins < 1 Then
    If Timer1.Interval > 0 Then
      Call SendPing
      Call CheckIpNotify
    End If
  Else
    iElapsedMin = iElapsedMin + 1
   ' Timer1.Enabled = (iElapsedMin < vMins)
    If iElapsedMin = vMins Then
      Call SendPing
      Call CheckIpNotify
      iElapsedMin = 0
    End If
  End If
End Sub

Public Sub ResetNextCounter()
  lblNext = UserSettings.Interval \ 1000
End Sub

Private Sub tmNext_Timer()
  lblNext = Val(lblNext) - 1
End Sub

Private Sub CheckIpNotify()
  'put this in SendPing to always send notifications, for my purposes, I only
  'want it to send email when the timer fires and finds a bad ip
  Call ResetNextCounter
  If lstBadIP.ListCount > 0 Then
    If UserSettings.Notify Then
      If UserSettings.Retry Then
        tmRetry.Interval = UserSettings.RetryIntvl
        tmRetry.Enabled = True
      Else
        Call SendEmail
      End If
    End If
  End If
  
End Sub

Private Sub tmRetry_Timer()
  Static tries As Integer
  
  Call SendPing 'retry ip list to refresh bad/good list
  With tmRetry
    If lstBadIP.ListCount = 0 Then
      .Enabled = False
      tries = 0
      Exit Sub
    Else
      tries = tries + 1
      If tries >= UserSettings.RetryNum Then
        SendEmail
        tries = 0
        .Enabled = False
        Exit Sub
      End If
    End If
    .Enabled = False
    .Enabled = True 'reset retry timer and keep trying until tries=RetryNum or no more bad IPs
  End With
End Sub

Private Sub SendPing()
  Dim ping As PingState
  Dim intX As Integer
  Dim intC As Integer
  
  intC = UserSettings.ListCount
  If intC > 0 Then
    lblNow = Now
    lbReturn.Clear
    lstGoodIP.Clear
    lstBadIP.Clear
    With UserSettings
      For intX = 0 To intC - 1
        If .Checked(intX) Then
          ping = PingIp(.List(intX))
          If ping = Disconnected Then
            lstBadIP.AddItem .List(intX)
          Else
            lstGoodIP.AddItem .List(intX)
          End If
        End If
      Next
      If .KeepLog Then Call AppendToLog
    End With
  End If
End Sub

Private Sub AppendToLog()
  Dim intX As Integer
  On Error GoTo Exit_Here
  
  Open UserSettings.LogName For Append As #1
    Print #1, vbCrLf & String(15, "-") & lblNow & String(15, "-")
    With lbReturn
      For intX = 0 To .ListCount - 1
        Print #1, .List(intX)
      Next
    End With
  Close
  
Exit_Here:
  cmdLog.Enabled = (Len(Dir(UserSettings.LogName)) > 1)
  Exit Sub
End Sub

Private Sub cmdPing_Click()
  Call SendPing
End Sub

Private Function GetAppVer() As String
  With App
    GetAppVer = " v" & .Major & "." & .Minor & " "
  End With
End Function

Private Sub SendEmail()
  Dim msg As String
  Dim intX As Integer
  
  Screen.MousePointer = vbHourglass
  
  msg = "This is an automated response from the Ping Monitor utility!" & vbCrLf & vbCrLf
  msg = msg & "The following IP's are off the network:" & vbCrLf
  With lstBadIP
    For intX = 0 To .ListCount - 1
      msg = msg & .List(intX) & vbCrLf
    Next
  End With
  
'    .username = "JJones" ' network user name and password !
'    .Password = "******"
  mapiLogOn.SignOn ' use current user
  
  Do While mapiLogOn.SessionID = 0
    DoEvents ' need to wait until the new session is created
  Loop
  
  With UserSettings
    For intX = 0 To .EmailCount - 1
      Call SendToEmail(.Email(intX), msg)
    Next
  End With
  mapiLogOn.SignOff
  
  Screen.MousePointer = vbNormal
End Sub
Public Sub SendToEmail(ByVal Email As String, ByVal msg As String)
  With MAPIMessages1
    'create a new message and address it
    .SessionID = mapiLogOn.SessionID
    .Compose
    .RecipDisplayName = Email
    .AddressResolveUI = True
    .RecipAddress = "smtp:" & Email
    
    .MsgSubject = "NETWORK PROBLEM DETECTED"
    .MsgNoteText = msg
    .Send False
  End With
End Sub
