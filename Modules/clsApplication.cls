VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get Version() As String
  Version = myVersion
End Property


Public Sub SendLine(ByVal lIndex As Long, ByVal CmdLine As String)
  On Error Resume Next
  If lIndex < 0 Or lIndex >= c_MaxUsers Then Exit Sub
  o_Clients(lIndex).SendLine CmdLine
  If Err Then Err.Clear
End Sub


Public Sub Disconnect(ByVal Index As Long)
  On Error Resume Next
  o_Clients(Index).WSU.Close
  If Err Then Err.Clear
End Sub


Public Property Get HostIP(Optional ByVal Index As Long = -1) As String
  On Error Resume Next
  If Index < 0 Then
    HostIP = frmMain.WS.LocalIP
  Else
    HostIP = o_Clients(Index).WSU.RemoteHostIP
  End If
  If Err Then Err.Clear
End Property


Public Property Get HostName(Optional ByVal Index As Long = -1) As String
  On Error Resume Next
  If Index < 0 Then
    HostName = frmMain.WS.LocalHostName
  Else
    HostName = o_Clients(Index).WSU.RemoteHost
  End If
  If Err Then Err.Clear
End Property


Public Property Get UpTime(Optional ByVal Index As Long = -1) As Long
  On Error Resume Next
  If Index < 0 Then
    UpTime = SecondsLeftAfter(m_ServerUpTime)
  Else
    UpTime = SecondsLeftAfter(o_Clients(Index).ConnTime)
  End If
  If UpTime < 0 Then UpTime = 0
  If Err Then Err.Clear
End Property

