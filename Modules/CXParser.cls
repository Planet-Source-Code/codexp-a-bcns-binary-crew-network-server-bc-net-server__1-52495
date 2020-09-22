VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCXParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------'
' clsCXParser Class Copyright (C)2003 by CodeXP                   CodeXP@Lycos.de '
'---------------------------------------------------------------------------------'
' YOU CAN USE, YOU CAN CHANGE IT, BUT GIVE IT OTHERS AS IS (WITHOUT CHANGES)      '
'---------------------------------------------------------------------------------'
Option Explicit

Private Const c_MaxParams = 255           ' Maximum Count of Params '

' CommandLine Variables       '
Private m_CommandLine         As String   ' CommandLine             '
Private m_Command             As String   ' Full Command            '
Private m_Cmd                 As String   ' Command only            '
Private m_Class               As String   ' Command Class           '
Private m_User                As String   ' Full User ID            '
Private m_UserID              As String   ' User ID only            '
Private m_Server              As String   ' Server                  '
Private m_Params              As String   ' All Params              '
Private m_ParamCount          As Long     ' Count of Params         '
Private m_Param(c_MaxParams)  As String   ' Param Array             '
Private m_Message             As String   ' Message                 '

Public myClassName            As String   'Class Name of this Parser'

Public Event Exec(Cmd As clsCXParser, ByRef bExecuted As Boolean)


'================================================='
' Commandline Syntax:                             '
' "[:user ]cmd[ params][:message]"                '
' Specification:                                  '
' Commandline can contain four parts:             '
' UserID, Command, Parameters and Message.        '
' The Command part must be present in all cases,  '
' all other parts are optional and may be missing.'
' The parts in Commandline are separated by one   '
' or more White Spaces.(White Space: just chr(32))'
'-------------------------------------------------'
' 1. Part: "[:user ]" (Optional)                  '
'    Description: User ID                         '
'    >user: UserID[@Server]                       '
'                                                 '
' 2. Part: "cmd" (Not Optional)                   '
'    Description: Command                         '
'    >cmd: [Class.]Command                        '
'                                                 '
' 3. Part: "[ params]" (Optional)                 '
'    Description: Parameters are used to extend   '
'                 the Command.                    '
'    >params: Param1[ Param2][ Param3][ ParamX]...'
'                                                 '
' 4. Part: "[:message]" (Optional)                '
'    Description: Message or special String that  '
'                 can contain all Ascii characters'
'                 excepted NewLine(CrLf).         '
'================================================='


Private Sub ParseCMDL()
  Dim n_CommandLine         As String   ' CommandLine             '
  Dim n_Command             As String   ' Full Command            '
  Dim n_Cmd                 As String   ' Command only            '
  Dim n_Class               As String   ' Command Class           '
  Dim n_User                As String   ' Full User ID            '
  Dim n_UserID              As String   ' User ID only            '
  Dim n_Server              As String   ' Server                  '
  Dim n_Params              As String   ' All Params              '
  Dim n_ParamCount          As Long     ' Count of Params         '
  Dim n_Message             As String   ' Message                 '
  
  Dim CTokens               As New clsTokens
  Dim i                     As Long
  
  n_CommandLine = LTrim(m_CommandLine)
  
  ' Get and Delete Message            '
  i = 1
  If Left(n_CommandLine, 1) = ":" Then
    i = InStr(n_CommandLine, ":") + 1
  End If
  i = InStr(i, n_CommandLine, ":")
  If i > 0 Then
    n_Message = Mid(n_CommandLine, i + 1)
    n_CommandLine = Left(n_CommandLine, i - 1)
  End If
  
  CTokens.Init n_CommandLine
  
  ' Get and Delete User               '
  If CTokens.Count Then
    If Left(CTokens.Token(1), 1) = ":" Then
      n_User = Mid(CTokens.Token(1), 2)
      CTokens.Remove 1
    End If
  End If
  
  ' Get and Delete Command            '
  If CTokens.Count Then
    n_Command = CTokens.Token(1)
    CTokens.Remove 1
  End If

  ' Get all Params                    '
  n_Params = CTokens.Data
  n_ParamCount = IIf(CTokens.Count > c_MaxParams, c_MaxParams, CTokens.Count)
  For i = 1 To n_ParamCount
    m_Param(i) = CTokens.Token(i)
  Next i

  ' Split Command                     '
  If Len(n_Command) > 0 Then
    CTokens.Init n_Command, ".", 2, True
    If CTokens.Count = 1 Then
      n_Cmd = CTokens.Token(1)
    Else
      n_Class = CTokens.Token(1)
      n_Cmd = CTokens.Token(2)
    End If
  End If
  
  ' Split User                        '
  If Len(n_User) > 0 Then
    CTokens.Init n_User, "@", 2, True
    If CTokens.Count = 1 Then
      n_UserID = CTokens.Token(1)
    Else
      n_UserID = CTokens.Token(1)
      n_Server = CTokens.Token(2)
    End If
  End If
  
  ' Update Data       '
  m_Command = n_Command
  m_Cmd = n_Cmd
  m_Class = n_Class
  m_User = n_User
  m_UserID = n_UserID
  m_Server = n_Server
  m_Params = n_Params
  m_ParamCount = n_ParamCount
  m_Message = n_Message
End Sub


Public Function Parse(ByVal CommandLine As String) As Boolean
  Dim CTokens As New clsTokens
  Dim bExec   As Boolean
  
  m_CommandLine = CommandLine
  CTokens.Init CommandLine, vbCrLf
  If CTokens.Count > 0 Then m_CommandLine = CTokens.Token(1)
  ParseCMDL
  If (m_Class = myClassName) Or myClassName = "{ALL}" Then RaiseEvent Exec(Me, bExec)
  
  Parse = bExec
End Function


Public Property Get CommandLine() As String         ' CommandLine             '
  CommandLine = m_CommandLine
End Property


Public Property Get Command() As String             ' Full Command            '
  Command = m_Command
End Property


Public Property Get Cmd() As String                 ' Command only            '
  Cmd = m_Cmd
End Property


Public Property Get Class() As String               ' Command Class           '
  Class = m_Class
End Property


Public Property Get User() As String                ' Full User ID            '
  User = m_User
End Property


Public Property Get UserID() As String              ' User ID only            '
  UserID = m_UserID
End Property


Public Property Get Server() As String              ' Server                  '
  Server = m_Server
End Property


Public Property Get Params() As String              ' All Params              '
  Params = m_Params
End Property


Public Property Get ParamCount() As Long            ' Count of Params         '
  ParamCount = m_ParamCount
End Property


Public Property Get Param(ByVal Index As Long) As String     ' Param  '
  Param = m_Param(Index)
End Property


Public Property Get Message() As String             ' Message                 '
  Message = m_Message
End Property

