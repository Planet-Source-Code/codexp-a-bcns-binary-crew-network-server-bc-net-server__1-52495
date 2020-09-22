VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const c_ModuleName = "Comm"

Private SAI         As clsNetAppInterface ' Server Application Interface  '
Private o_Users     As clsUsers           ' User Collection               '
Private m_bStarted  As Boolean            ' Module Started (True/False)   '


' Interface Property: Started()                                                         '
Public Property Get Started() As Boolean
  ' Return the State of Module '
  Started = m_bStarted
End Property


' Interface Property: ModuleName()                                                      '
Public Property Get ModuleName() As String
  ' Return the Name of this Module '
  ModuleName = c_ModuleName
End Property


' Interface Property: ModuleVersion()                                                   '
Public Function ModuleVersion() As String
  ' Return Version of this Module '
  ModuleVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function


' Interface Event: Connected()                                                          '
Public Sub Connected(ByVal Index As Long)
  ' Event raises on connecting Socket(Index) '
End Sub


' Interface Event: Disconnected()                                                       '
Public Sub Disconnected(ByVal Index As Long)
  ' Event raises on closing Socket(Index) '
  Dim sUsr As String
  sUsr = o_Users.GetUserID(Index)
  If o_Users.UserExists(sUsr) Then
    o_Users.Remove sUsr
    o_Users.SendToAll ":" & sUsr & "@" & SAI.HostName & " Comm.Leave:Disconnected"
  End If
End Sub


' Interface Function: StartModule()                                                     '
Public Function StartModule(oNetApp As Object) As Boolean
  ' Here we have to Start our Module.   '
  '> Return True if Module was started, '
  '  else Return False if not.          '
  '> We have to set a Reference to the  '
  '  Server Application Interface       '
  '  (oNetApp).                         '
  
  On Error GoTo StartModule_Resume
  
  ' If already started then Exit: '
  If m_bStarted Then Exit Function
  
  ' Set Reference to the Server Application Interface '
  If Not SAI.Init(oNetApp) Then Exit Function
  
  ' Check Server Version  '
  If Val(SAI.Version) < Val(ModuleVersion) Then
    ' Wrong Server Version! '
  Else
    ' Return True: '
    StartModule = True
  End If

StartModule_Resume:
  m_bStarted = StartModule
End Function


' Interface Event: Execute()                                                            '
Public Function Execute(ByVal Index As Long, ByVal CmdLine As String) As Boolean
  ' Event raises on DataArrival on Socket(Index)    '
  '> Here we have to Parse and Execute the          '
  '  Command Line (CmdLine).                        '
  '> We can send an answer with SendLine() Function.'
  '> Return False if Command is unknown,            '
  '  else Return True if Command is known.          '
  
  Dim oP  As New clsCXParser
  Dim Tmp As String
  Dim i   As Long
  
  ' Parse Command Line: '
  oP.Parse CmdLine
  
  ' If Module is not started then Exit: '
  If Not m_bStarted Then Exit Function
  ' If Class does not match ModuleName then Exit: '
  If (UCase(oP.Class) <> UCase(ModuleName)) Then Exit Function
  
  Execute = True
  ' Execute Command: '
  Select Case UCase(oP.Cmd)
  '>Command: "PrivMsg"      '
  Case "PRIVMSG"
  
    Dim sUsr As String
    sUsr = o_Users.GetUserID(Index)
    If Not o_Users.UserExists(sUsr) Then sUsr = ""
    
    If sUsr = "" Then
      SAI.SendLine Index, "Comm.Error PrivMsg:Access Denied"
    Else
      If oP.ParamCount > 0 Then
        Dim sTo As String
        sTo = oP.Param(1)
        If o_Users.UserExists(sTo) Then
          sTo = o_Users(sTo).UserID
          Dim sParams As String
          For i = 2 To oP.ParamCount
            If Len(sParams) > 0 Then sParams = sParams & " "
            sParams = sParams & oP.Param(i)
          Next i
          Tmp = ":" & sUsr & " Comm.PrivMsg " & sTo & " " & sParams & ":" & oP.Message
          o_Users(sUsr).SendLine Tmp
          o_Users(oP.Param(1)).SendLine Tmp
        Else
          SAI.SendLine Index, "Comm.Error PrivMsg:UserID does not exists"
        End If
      Else
        SAI.SendLine Index, "Comm.Error PrivMsg:UserID is required"
      End If
    End If
    
  '>Command: "Login"        '
  Case "LOGIN"
    
    If oP.ParamCount > 1 Then
      Tmp = o_Users.GetUserID(Index)
      If Len(Tmp) > 0 Then
        SAI.SendLine Index, "Comm.Error Login:you are already logged in"
      Else
        If o_Users.UserExists(oP.Param(1)) Then
          SAI.SendLine Index, "Comm.Error Login:Login rejected, UserID already in use"
        Else
          Tmp = oP.Param(1)
          If o_Users.LoginUser(Index, Tmp, oP.Param(2)) Then
            Tmp = o_Users(Tmp).UserID
            ' Send Login Ok to User         '
            SAI.SendLine Index, ":" & Tmp & "@" & SAI.HostName & " Comm.Login OK:Login successful"
            ' Send User List to User        '
            o_Users.SendUserList Index, Index
            ' Send User Login to Others     '
            o_Users.SendToAll ":" & Tmp & "@" & SAI.HostName & " Comm.Login", Index
            ' Send Welcome Message to User  '
            SendWelcomeMsg Index
          Else
            SAI.SendLine Index, "Comm.Error Login:Login rejected, check your UserID and Password"
          End If
        End If
      End If
    Else
      SAI.SendLine Index, "Comm.Error Login:UserID and Password are required"
    End If
  
  '>Command: "Leave"        '
  Case "LEAVE"
  
    If o_Users.UserExists(o_Users.GetUserID(Index)) Then
      o_Users.Remove o_Users.GetUserID(Index)
      SAI.Disconnect Index
    Else
      SAI.SendLine Index, "Comm.Error Leave:Access Denied"
    End If
    
  '>Unknown Command         '
  Case Else
  
    ' Return False: '
    Execute = False
    
  End Select
  
End Function


Private Sub SendSrvMsg(ByVal Index As Long, ByVal SrvMsg As String, Optional ByVal sParams As String)
  SAI.SendLine Index, ":Server@" & SAI.HostName & " Comm.SrvMsg " & sParams & ":" & SrvMsg
End Sub


Private Sub SendWelcomeMsg(ByVal Index As Long)
  Dim Tmp As String
  Tmp = o_Users.GetUserID(Index)
  If Len(Tmp) > 0 Then
    SendSrvMsg Index, "Hello " & Tmp & "!", "Info Begin"
    SendSrvMsg Index, "Welcome to BCC Server on " & SAI.HostName & "!", "Info"
    SendSrvMsg Index, "BCC Server Module Version " & ModuleVersion, "Info"
    SendSrvMsg Index, "Copyright (C)2003 by CodeXP", "Info"
    SendSrvMsg Index, "", "Info"
    SendSrvMsg Index, "General Server Statistics:", "Info"
    SendSrvMsg Index, "Users online: " & o_Users.UsersCount, "Info"
    SendSrvMsg Index, "", "Info End"
  End If
End Sub


Private Sub Class_Initialize()
  Set SAI = New clsNetAppInterface
  Set o_Users = New clsUsers
  Call o_Users.Init(SAI)
End Sub


Private Sub Class_Terminate()
  If Started Then
    Dim cUsr As clsUser
    On Error Resume Next
    For Each cUsr In o_Users
      Call SAI.Disconnect(cUsr.Index)
      o_Users.Remove cUsr.UserID
    Next cUsr
  End If
  Set o_Users = Nothing
  Set SAI = Nothing
End Sub
