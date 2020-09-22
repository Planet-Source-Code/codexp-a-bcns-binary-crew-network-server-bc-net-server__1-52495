Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------'
' modMain (BCNetServer Module) Copyright (C)2004 by CodeXP        CodeXP@Lycos.de '
'---------------------------------------------------------------------------------'
' YOU CAN USE, YOU CAN CHANGE IT, BUT GIVE IT OTHERS AS IS (WITHOUT CHANGES)      '
'---------------------------------------------------------------------------------'
Option Explicit

Public Const c_MaxUsers = 10
Public Const c_NetPort = 105

Public o_Clients(0 To c_MaxUsers - 1) As frmClient
Public o_Application  As clsApplication
Public o_Modules      As New clsModules
Public m_StartServer  As Boolean
Public m_DebugWindow  As Boolean
Public m_AppUpTime    As Double
Public m_ClientCount  As Long
Public m_FreeCount    As Long
Public m_ServerUpTime As Double
Public m_DataIn       As Long
Public m_DataOut      As Long
Public m_PFaults      As New Collection


Public Sub Main()
  Load frmMain
  frmMain.Show
End Sub


Public Function GetPFaults(ByVal IP As String) As Long
  If PFaultExists(IP) Then GetPFaults = m_PFaults(IP)
End Function


Public Sub AddPFault(ByVal sType As String, ByVal IP As String)
  Dim i As Long
  
  If PFaultExists(IP) Then
    i = m_PFaults(IP)
    m_PFaults.Remove IP
  End If
  
  i = i + 1
  m_PFaults.Add i, IP
  
  If i = 10 Then
    Debug.Print "############# Client Banned (" & IP & ")"
    LogEvent "Client Banned (" & IP & ")"
  End If
End Sub


Private Function PFaultExists(ByVal sKey As String) As Boolean
  Dim i As Long
  On Error Resume Next
  i = m_PFaults(sKey)
  If Err Then
    Err.Clear
  Else
    PFaultExists = True
  End If
End Function


Public Sub SendToAll(ByVal CmdLine As String, Optional ByVal ExceptIndex As Long = -1)
  Dim i As Long
  For i = 0 To c_MaxUsers - 1
    If i <> ExceptIndex Then
      If Not o_Clients(i) Is Nothing Then
        If o_Clients(i).WSU.State = 7 Then
          o_Clients(i).SendLine CmdLine
        End If
      End If
    End If
  Next i
End Sub


Public Function GetSocketState(ByVal State As Long) As String
  Dim Ret As String
  Select Case State
  Case 0:    Ret = "Closed"
  Case 1:    Ret = "Open"
  Case 2:    Ret = "Listening"
  Case 3:    Ret = "Connection Pending"
  Case 4:    Ret = "Resolving Host"
  Case 5:    Ret = "Host Resolved"
  Case 6:    Ret = "Connecting"
  Case 7:    Ret = "Connected"
  Case 8:    Ret = "Closing"
  Case 9:    Ret = "Error"
  Case Else: Ret = "Unknown State"
  End Select
  GetSocketState = Ret
End Function


Public Sub LogEvent(ByVal Msg As String)
  frmDebug.LogMessage Msg
  If m_DebugWindow Then
    If Not frmDebug.Visible Then frmDebug.Show
  End If
End Sub


Public Sub LogCommand(ByVal Msg As String)
  frmDebug.LogCommand Msg
  If m_DebugWindow Then
    If Not frmDebug.Visible Then frmDebug.Show
  End If
End Sub


Public Sub LogError(ByVal Msg As String)
  frmDebug.LogError Msg
  If m_DebugWindow Then
    If Not frmDebug.Visible Then frmDebug.Show
  End If
End Sub


Public Function myVersion() As String
  myVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function


Public Function myPath() As String
  myPath = AddBSlash(App.Path)
End Function


Public Function AddBSlash(ByVal fPath As String) As String
  AddBSlash = IIf(Right(fPath, 1) = "\", fPath, fPath & "\")
End Function


Public Function BytesView(ByVal Bytes As Long) As String
  Dim Ret As String
  Dim sME As String
  Dim dBy As Double
  
  sME = " Bytes"
  dBy = Bytes
  
  If dBy > 1024 Then
    sME = " KB"
    dBy = dBy / 1024
    If dBy > 1024 Then
      sME = " MB"
      dBy = dBy / 1024
      If dBy > 1024 Then
        sME = " GB"
        dBy = dBy / 1024
      End If
    End If
  End If
  
  BytesView = Val(Format(dBy, "0.00")) & sME
End Function


Public Function LngMsgBox(ByVal lMessage As Long) As VbMsgBoxResult
  Dim sMessage  As String
  Dim sTitle    As String
  Dim lStyle    As VbMsgBoxStyle
  
  Select Case lMessage
  Case 0  ' TEST MESSAGE '
    sMessage = "This is a Test Message with a (i) Logo, " & _
               "Title like ""MessageBox"" " & _
               "and [Yes] / [No] Buttons!"
    sTitle = "MessageBox"
    lStyle = vbInformation Or vbYesNo
  Case 100
    sMessage = "Can not connect to Server!"
    sTitle = "Connection Error!"
    lStyle = vbCritical
  Case 101
    sMessage = "No more free Sockets to connect to Server!"
    sTitle = "Server Full!"
    lStyle = vbExclamation
  Case 1001
    sMessage = "Local Client is not yet implemented!"
    sTitle = "Not implemented!"
    lStyle = vbExclamation
  Case Else
    sMessage = "Required Message is not set!"
    sTitle = "Error:"
    lStyle = vbCritical
    Exit Function
  End Select
  
  LngMsgBox = MsgBox(sMessage, lStyle, sTitle)
End Function


Public Sub DoMoreEvents(Optional ByVal HowMuch As Long = 100)
  Dim i As Long
  For i = 1 To HowMuch: DoEvents: Next i
End Sub


' Function ModuleName() - Returns Name of Module from Filename as String    '
Public Function ModuleName(ByVal sPath As String) As String
  sPath = Replace(sPath, "/", "\")
  If InStr(sPath, "\") > 0 Then sPath = Mid(sPath, InStrRev(sPath, "\") + 1)
  If InStr(sPath, ".") > 0 Then sPath = Left(sPath, InStrRev(sPath, ".") - 1)
  ModuleName = sPath
End Function


' Function PPL() - Returns incremental Value of the first Parameter as Long '
Public Function PPL(ByRef lInc As Long) As Long
  lInc = lInc + 1
  PPL = lInc
End Function


' Function MML() - Returns decremental Value of the first Parameter as Long '
Public Function MML(ByRef lDec As Long) As Long
  lDec = lDec - 1
  MML = lDec
End Function


Public Function RemDouble(ByVal strText As String, Optional ByVal sChar As String = " ") As String
  While InStr(strText, String(2, sChar))
    strText = Replace(strText, String(2, sChar), String(1, sChar))
  Wend
  RemDouble = strText
End Function
