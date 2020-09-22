VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetAppInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private o_NetApp  As Object ' Reference to the Server Application Interface '


' Public Function: Init()                                                               '
Public Function Init(ByVal oNetApp As Object) As Boolean
  If Not (oNetApp Is Nothing) Then
    Init = True
    ' Set a Reference to Server Application Interface: '
    Set o_NetApp = oNetApp
  End If
End Function


' Interface Action: Disconnect()                                                        '
Public Sub Disconnect(ByVal Index As Long)
  On Error Resume Next
  ' Close Winsock(Index): '
  Call o_NetApp.Disconnect(Index)
End Sub


' Interface Action: SendLine()                                                          '
Public Function SendLine(ByVal Index As Long, ByVal CmdLine As String)
  On Error Resume Next
  ' Send Command Line to Socket(Index): '
  Call o_NetApp.SendLine(Index, CmdLine)
End Function


' Interface Property: Version()                                                         '
Public Function Version() As String
  On Error Resume Next
  ' Return Server Application's Version: '
  Version = o_NetApp.Version()
End Function


' Interface Property: UpTime()                                                          '
Public Function UpTime(Optional ByVal Index As Long = -1) As Long
  On Error Resume Next
  If Index < 0 Then
    ' Return Server's UpTime: '
    UpTime = o_NetApp.UpTime()
  Else
    ' Return Client's UpTime: '
    UpTime = o_NetApp.UpTime(Index)
  End If
End Function


' Interface Property: HostName()                                                        '
Public Function HostName(Optional ByVal Index As Long = -1) As String
  On Error Resume Next
  If Index < 0 Then
    ' Return Server's HostName: '
    HostName = o_NetApp.HostName()
  Else
    ' Return Client's HostName: '
    HostName = o_NetApp.HostName(Index)
  End If
End Function


' Interface Property: HostIP()                                                          '
Public Function HostIP(Optional ByVal Index As Long = -1) As String
  On Error Resume Next
  If Index < 0 Then
    ' Return Server's HostIP: '
    HostIP = o_NetApp.HostIP()
  Else
    ' Return Client's HostIP: '
    HostIP = o_NetApp.HostIP(Index)
  End If
End Function


