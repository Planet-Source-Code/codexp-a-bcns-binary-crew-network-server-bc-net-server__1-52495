VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Client"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SrvClient.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   1665
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Timer tmrWSU 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1080
         Top             =   0
      End
      Begin VB.Timer tmrClient 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   0
      End
      Begin MSWinsockLib.Winsock WSU 
         Left            =   600
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.ListBox lstOut 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1710
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Index As Long

Public ConnTime As Double
Public bConnect As Boolean


Public Property Get Index() As Long
  Index = m_Index
End Property
'^'
Public Property Let Index(ByVal Value As Long)
  m_Index = Value
  ResetClient
  tmrClient.Enabled = True
  Me.Left = GetLeftByIndex(Index)
End Property


Private Function GetLeftByIndex(ByVal Index As Long) As Single
  Dim lWPSH   As Long ' Windows Per Screen H  '
  Dim lWPSV   As Long ' Windows Per Screen V  '
  lWPSV = Int(Screen.Height / Height)
  lWPSH = Int(Screen.Width / Width)
  Index = Index Mod lWPSV
End Function


Private Sub Form_Load()
  tmrWSU.Enabled = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Cancel = True
    Me.Hide
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  tmrClient.Enabled = False
  tmrWSU.Enabled = False
  WSU.Close
  DoMoreEvents
  tmrWSU_Timer
  DoMoreEvents
End Sub


Public Sub ResetClient()
  lstOut.Clear
  ConnTime = 0
  bConnect = False
  Me.Caption = "Client"
End Sub


Public Sub ExecCmd(ByVal CmdLine As String)
  Static lEPSC  As Long
  Static lEPST  As Long
  Dim Tmp As String
  Dim oP  As New clsCXParser
  Dim i   As Long
  
  If ConnTime = 0 Then tmrWSU_Timer
  
  oP.Parse CmdLine
  AddMsg Format(m_Index, "\[000\] ") & ">>> " & CmdLine
  
  Select Case UCase(oP.Class)
  Case "", "{ALL}"
    Select Case UCase(oP.Cmd)
    Case "VERSION"
      
      If oP.ParamCount > 0 Then
        ' GOT VERSION OF OTHER SIDE '
      Else
        SendLine "Version " & myVersion
      End If
      
    Case "HELP"
      
      SendLine "Error Help:Help is not yet Implemented"
    
    Case "MODULES"
    
      Dim bListModules As Boolean
      
      If oP.ParamCount > 0 Then
        
        Select Case UCase(oP.Param(1))
        Case "LIST"
          
          bListModules = True
          
        Case "LOAD"
        
          If oP.ParamCount > 1 Then
            Tmp = oP.Param(2)
            If o_Modules.ModuleExists(Tmp) Then
              SendLine "Error Modules Load:Module is already loaded"
            Else
              If o_Modules.ReadModules(Tmp) <> 0 Then
                Tmp = o_Modules(Tmp).ModuleName
                If o_Modules(Tmp).StartModule Then
                  SendLine "Modules Load " & Tmp & ":Module successful loaded"
                  SendToAll "Modules Load " & Tmp, m_Index
                Else
                  SendLine "Error Modules Load " & Tmp & ":Module can not be started"
                End If
              Else
                SendLine "Error Modules Load:Module can not be loaded"
              End If
            End If
          Else
            SendLine "Error Modules Load:Module Name is required"
          End If
        
        Case "UNLOAD"
        
          If oP.ParamCount > 1 Then
            Tmp = oP.Param(2)
            If o_Modules.ModuleExists(Tmp) Then
              o_Modules.RemoveModule oP.Param(2)
              SendLine "Modules Unload " & Tmp & ":Module successful removed"
              SendToAll "Modules Unload " & Tmp, m_Index
            Else
              SendLine "Error Modules Unload:Module does not exists"
            End If
          Else
            SendLine "Error Modules Unload:Module Name is required"
          End If
        
        Case Else
        
          SendLine "Error Modules:Unknown Parameter"
          
        End Select
      
      Else
        
        bListModules = True
      
      End If
      
      If bListModules Then
      
        If o_Modules.ModulesCount > 0 Then
          
          Dim oMod  As clsModule
          Dim sML   As String
          
          For Each oMod In o_Modules
            If Len(sML) > 0 Then sML = sML & ";"
            If Len(sML) = 0 Then sML = "Modules List:"
            sML = sML & oMod.ModuleName
          Next oMod
          
          SendLine sML
        
        Else
        
          SendLine "Modules None"
        
        End If
        
      End If
      
    Case "ERROR"
      
      ' DO NOTHING (OR HANDLE ERRORS) '
      lEPSC = lEPSC + 1
    
    Case Else
      
      If Len(oP.Cmd) > 0 Then
        lEPSC = lEPSC + 1
        SendLine "Error " & oP.Command & ":Unknown Command"
      End If
      
    End Select
    
  Case Else ' other Classes (Modules) '
  
    If o_Modules.ModuleExists(oP.Class) Then
      If o_Modules(oP.Class).Started Then
        If Not o_Modules(oP.Class).Execute(m_Index, CmdLine) Then
          
          lEPSC = lEPSC + 1
          SendLine "Error " & oP.Command & ":Unknown Command"
        
        End If
      Else
          
          lEPSC = lEPSC + 1
          SendLine "Error Module Disabled " & oP.Class & ":Module is presently disabled"
      
      End If
    Else
      
      lEPSC = lEPSC + 1
      SendLine "Error Module Missing " & oP.Class & ":Module is not implemented"
    
    End If
    
  End Select
  
  i = TickMilli - lEPST
  If (i <= 10) And (lEPSC >= 10) Then
    ProtectionFault "Overrun"
  ElseIf i >= 10 Then
    lEPSC = 0
  End If
  lEPST = TickMilli
End Sub


Public Sub AddMsg(ByVal sMsg As String)
  Dim sT()  As String
  Dim i     As Long
  
  sT = Split(sMsg, vbCrLf)
  For i = 0 To UBound(sT)
    If (i = UBound(sT)) And (sT(i) = "") And (i > 1) Then Exit For
    lstOut.AddItem IIf(i = 1, TimeStamp, "") & sT(i)
  Next i
  
  While lstOut.ListCount > 1000
    lstOut.RemoveItem 0
  Wend
  
  If lstOut.ListCount > 0 Then
    lstOut.TopIndex = lstOut.ListCount - 1
  End If
End Sub


Public Sub SendLine(ByVal CmdLine As String)
  ' Check if we are connected '
  If WSU.State = 7 Then
    ' Count up sent Data Bytes    '
    m_DataOut = m_DataOut + Len(CmdLine) + 2
    ' Send Data & vbCrLf          '
    WSU.SendData CmdLine & vbCrLf
    DoMoreEvents  ' System Events '
    AddMsg Format(m_Index, "\[000\] ") & "<<< " & CmdLine
  Else
    ' We are not Connected: Log Error '
    LogError "Error: " & Format(m_Index, "\[000\] ") & "Can't send because not connected"
  End If
End Sub


Private Sub tmrClient_Timer()
  'Dim bVis As Boolean
  'bVis = (WSU.State = 7)
  'If Me.Visible <> bVis Then Me.Visible = bVis
End Sub


Private Sub tmrWSU_Timer()
  ' Switch Socket Status  '
  Select Case WSU.State
  ' Socket is Connecting or is Resolving Host                             '
  Case 6, 4
  
    ' do nothing  '
    
  ' Socket is Connected                                                   '
  Case 7
  
    bConnect = False
    ' Check if we were connected  '
    If ConnTime = 0 Then
      ' Reset Client and Init some Variables    '
      ResetClient
      ' Initialize Variables on new Connection  '
      ConnTime = TimeLong
      ' Log connected Event                     '
      LogEvent Format(m_Index, "\[000\] ") & WSU.RemoteHostIP & " Connected"
      ' Raise Event Connected on Modules        '
      o_Modules.Connected Index
    End If
    
  ' Socket is Closed                                                      '
  Case 0
  
    ' Check if we was connected  '
    If ConnTime <> 0 Then
      ' Log disconnect Event                     '
      LogEvent Format(m_Index, "\[000\] ") & WSU.RemoteHostIP & " Closed"
      WSU_DataArrival -1  ' Clear Buffer  '
      ConnTime = 0
      ' Raise Event Disconnected on Modules     '
      o_Modules.Disconnected Index
    End If
    
    ' If we try to Connect to Server  '
    If bConnect Then
      bConnect = False
      ' Error occured on connecting  '
      LngMsgBox 100
    End If
  
  ' All Other States                                                      '
  Case Else
  
    ' Close Socket  '
    WSU.Close
    
  End Select
End Sub


Private Sub WSU_DataArrival(ByVal bytesTotal As Long)
  Static lInstances As Long
  Static lFlood     As Long
  Static lRTime     As Long
  Static Buffer     As String
  Dim oT  As New clsTokens
  Dim Dop As String
  Dim Tmp As String
  Dim i   As Long
  
  On Error Resume Next
  
  If bytesTotal < 0 Then
    ' Reset Buffer  '
    Buffer = ""
    Exit Sub
  End If
  
  ' Count up recursive Instances  '
  lInstances = lInstances + 1
  If lInstances > 100 Then
    ProtectionFault "Overflow"
    Exit Sub
  End If
  
  ' Count up recieved Data Bytes  '
  m_DataIn = m_DataIn + bytesTotal
  
  ' Check Flooding  '
  i = TickMilli - lRTime
  If i >= 50 Then
    lFlood = 0
  ElseIf i >= 30 Then
    lFlood = lFlood - 10
  ElseIf i >= 10 Then
    lFlood = lFlood - 1
  ElseIf i < 5 Then
    lFlood = lFlood + 1
  End If
  If lFlood < 0 Then lFlood = 0
  
  ' Get recieved Data '
  WSU.GetData Tmp
  lRTime = TickMilli
  
  Dop = Tmp
  Dop = RemDouble(Dop, Chr(27))
  Dop = RemDouble(Dop)
  Dop = Replace(Dop, vbCrLf, "\n")
  Debug.Print "L(" & Len(Tmp) & ") """ & Dop & """"
'  Debug.Print "********** ::" & lFlood
  
  If Len(Buffer) Then
    ' Prepend Buffer on Data  '
    Tmp = Buffer & Tmp
    Buffer = ""
  End If
  
  ' Check Buffer Overflow   '
  If (Len(Buffer) > 255) Or (bytesTotal > 2500) Then
    ProtectionFault "Overflow"
    Exit Sub
  End If
  If lFlood > 20 Then
    ProtectionFault "Flood"
    Exit Sub
  End If
  
  ' Tokenize Data Lines     '
  oT.Init Tmp, vbCrLf, , True
  For i = 1 To oT.Count
    If i = oT.Count Then
      ' Save Data without vbCrLf in Buffer  '
      Buffer = oT.Token(i)
    Else
      ' Execute Command Line                '
      ExecCmd oT.Token(i)
    End If
  Next i
  
  ' Count down Instances    '
  lInstances = lInstances - 1
  DoMoreEvents
End Sub


Private Sub WSU_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  ' Log Error '
  LogError Format(m_Index, "\[000\] ") & "Socket Error: " & vbCrLf & Description
End Sub


Public Function SocketFree() As Boolean
  SocketFree = (ConnTime = 0) And (WSU.State = 0)
End Function


Public Function Connect(ByVal sHOST As String) As Boolean
  On Error Resume Next
  If WSU.State Then WSU.Close
  tmrWSU_Timer
  ' Try to connect to Host      '
  WSU.Connect sHOST, c_NetPort
  bConnect = True
  If Err = 0 Then Connect = True
End Function


Public Sub ProtectionFault(ByVal sType As String)
  Dim Msg As String
  Select Case UCase(Trim(sType))
  Case "FLOOD"
    Msg = "Flood Protection on Client " & m_Index
  Case "OVERRUN"
    Msg = "Errors Overrun on Client " & m_Index
  Case "OVERFLOW"
    Msg = "Buffer Overflow on Client " & m_Index
  End Select
  AddPFault Trim(sType), WSU.RemoteHostIP
  LogEvent Msg
  WSU.Close
End Sub
