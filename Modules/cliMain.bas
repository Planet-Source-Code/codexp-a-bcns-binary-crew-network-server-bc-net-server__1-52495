Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------'
' modMain (BC Console Module) Copyright (C)2004 by CodeXP         CodeXP@Lycos.de '
'---------------------------------------------------------------------------------'
' YOU CAN USE, YOU CAN CHANGE IT, BUT GIVE IT OTHERS AS IS (WITHOUT CHANGES)      '
'---------------------------------------------------------------------------------'
Option Explicit

Public Const c_MaxClients = 10
Public Const c_NetPort = 105
Public Const c_RegAppKey = "Software\CodeXP\BCConsole"

Public o_Cleints(0 To c_MaxClients - 1) As frmClient
Public m_DebugWindow  As Boolean
Public m_AppUpTime    As Double
Public m_DataIn       As Long
Public m_DataOut      As Long


Public Sub Main()
  Load frmMain
  frmMain.Show
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
'  frmDebug.LogMessage Msg
'  If m_DebugWindow Then
'    If Not frmDebug.Visible Then frmDebug.Show
'  End If
End Sub


Public Sub LogCommand(ByVal Msg As String)
'  frmDebug.LogCommand Msg
'  If m_DebugWindow Then
'    If Not frmDebug.Visible Then frmDebug.Show
'  End If
End Sub


Public Sub LogError(ByVal Msg As String)
  Debug.Print Msg
'  frmDebug.LogError Msg
'  If m_DebugWindow Then
'    If Not frmDebug.Visible Then frmDebug.Show
'  End If
End Sub


Public Function myVersion() As String
  myVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function


Public Function myPath() As String
  myPath = AddBSlash(App.Path)
End Function


Public Function AddBSlash(ByVal fPath As String) As String
  AddBSlash = IIf(Left(fPath, 1) = "\", fPath, fPath & "\")
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

