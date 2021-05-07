VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public EmailCount As Integer
Public Notify As Boolean
Private m_Email() As String

Public KeepLog As Boolean
Public LogName As String

Public Retry As Boolean
Public RetryNum As Integer
Public RetryIntvl As Long

Public Interval As Long
Public TTL As Integer
Public ListCount As Integer
Private m_List() As String
Private m_Checked() As Boolean

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Public Property Let List(ByVal index As Integer, ByVal sItem As String)
  m_List(index) = sItem
End Property

Public Property Get List(ByVal index As Integer) As String
  List = m_List(index)
End Property

Public Property Let Checked(ByVal index As Integer, ByVal pItem As Boolean)
  m_Checked(index) = pItem
End Property

Public Property Get Checked(ByVal index As Integer) As Boolean
  Checked = m_Checked(index)
End Property

Public Property Let Email(ByVal index As Integer, ByVal sItem As String)
  m_Email(index) = sItem
End Property

Public Property Get Email(ByVal index As Integer) As String
  Email = m_Email(index)
End Property
Public Sub RedimEmail()
  On Error Resume Next
  ReDim m_Email(0 To EmailCount - 1)
End Sub

Public Sub RedimList()
  On Error Resume Next
  ReDim m_List(0 To ListCount - 1)
End Sub

Public Sub RedimChecked()
  On Error Resume Next
  ReDim m_Checked(0 To ListCount - 1)
End Sub

Public Sub InitList(ByVal Count As Integer)
  ListCount = Count
  Call RedimList
  Call RedimChecked
End Sub

Public Sub InitEmail(ByVal Count As Integer)
  EmailCount = Count
  Call RedimEmail
End Sub

Public Sub SaveSettings()
  Dim intX As Integer
  Dim intL As Integer
  Dim intE As Integer
  Dim strTitle As String
  
  intL = ListCount
  intE = EmailCount
  
  strTitle = App.Title
  Call SaveSetting(strTitle, "Settings", "TTL", TTL)
  Call SaveSetting(strTitle, "Settings", "ListCount", intL)
  Call SaveSetting(strTitle, "Settings", "PingInterval", Interval)
  Call SaveSetting(strTitle, "Settings", "EmailCount", intE)
  Call SaveSetting(strTitle, "Settings", "Notify", Notify)
  Call SaveSetting(strTitle, "Settings", "Retry", Retry)
  Call SaveSetting(strTitle, "Settings", "RetryIntvl", RetryIntvl)
  Call SaveSetting(strTitle, "Settings", "RetryNum", RetryNum)
  Call SaveSetting(strTitle, "Settings", "LogName", LogName)
  Call SaveSetting(strTitle, "Settings", "KeepLog", KeepLog)
  
  If intL > 0 Then
    For intX = 0 To intL - 1
      Call SaveSetting(strTitle, "Settings", "List" & intX, List(intX))
      Call SaveSetting(strTitle, "Settings", "Checked" & intX, Checked(intX))
    Next
  End If
  
  If intE > 0 Then
    For intX = 0 To intE - 1
      Call SaveSetting(strTitle, "Settings", "Email" & intX, Email(intX))
    Next
  End If

End Sub
Public Sub GetSettings()
  Dim intX As Integer
  Dim intC As Integer
  Dim strTitle As String
  
  strTitle = App.Title
  TTL = GetSetting(strTitle, "Settings", "TTL", 5)
  Interval = GetSetting(strTitle, "Settings", "PingInterval", 3000)
  Notify = GetSetting(strTitle, "Settings", "Notify", False)
  Retry = GetSetting(strTitle, "Settings", "Retry", False)
  RetryIntvl = GetSetting(strTitle, "Settings", "RetryIntvl", 30000)
  RetryNum = GetSetting(strTitle, "Settings", "RetryNum", 2)
  
  LogName = GetSetting(strTitle, "Settings", "LogName", _
  BackSlash(App.Path) & GetName & ".txt")
  KeepLog = GetSetting(strTitle, "Settings", "KeepLog", False)
  
  intC = GetSetting(strTitle, "Settings", "ListCount", 0)
  ListCount = intC
  If intC > 0 Then
    Call RedimList
    Call RedimChecked
    
    For intX = 0 To intC - 1
      List(intX) = GetSetting(strTitle, "Settings", "List" & intX)
    Next
    For intX = 0 To intC - 1
      Checked(intX) = GetSetting(strTitle, "Settings", "Checked" & intX, "True")
    Next
  End If
  
  intC = GetSetting(strTitle, "Settings", "EmailCount", 0)
  EmailCount = intC
  If intC > 0 Then
    Call RedimEmail
    
    For intX = 0 To intC - 1
      Email(intX) = GetSetting(strTitle, "Settings", "Email" & intX)
    Next
  End If

End Sub
Private Function GetName() As String
  Dim username As String * 255
  
  Call GetUserName(username, 255)
  GetName = Left(username, InStr(username, Chr(0)) - 1)
  
End Function

