VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   AutoRedraw      =   -1  'True
   Caption         =   "BC Console"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Client.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   2445
   ScaleWidth      =   4755
   Visible         =   0   'False
   Begin VB.PictureBox picUserList 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   3360
      ScaleHeight     =   1695
      ScaleWidth      =   1215
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      Begin VB.ListBox lstUsers 
         Appearance      =   0  '2D
         Height          =   1020
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picSend 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4545
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtSend 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblCMDL 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   150
      End
   End
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
      Begin MSWinsockLib.Winsock WSU 
         Left            =   600
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
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
   End
   Begin VB.ListBox lstOut 
      Appearance      =   0  '2D
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
      Height          =   1740
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblWait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currently not Connected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   2070
   End
   Begin VB.Menu menuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuModules 
         Caption         =   "&Modules ..."
      End
      Begin VB.Menu mnuLine100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu menuCommands 
      Caption         =   "Commands"
      Visible         =   0   'False
      Begin VB.Menu mnuCmd 
         Caption         =   "Comm.Login %UserID% pass"
         Index           =   0
      End
      Begin VB.Menu mnuCmd 
         Caption         =   "Comm.Leave"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_LastID  As String
Private m_UserID  As String
Private m_SrvInfo As String
Private f_Modules As frmModules
Private o_PrivMsg As New Collection

Public ConnTime   As Double
Public bConnect   As Boolean
Public bLoaded    As Boolean


Public Property Get PrivMsgCount() As Long
  PrivMsgCount = o_PrivMsg.Count
End Property


Public Property Get PrivMsg(IndexKey) As frmPrivMsg
  Set PrivMsg = o_PrivMsg(IndexKey)
End Property


Public Function AddPrivMsg(ByVal sToUserID As String) As Boolean
  On Error GoTo AddPrivMsg_Resume
  
  sToUserID = Trim(sToUserID)
  
  If sToUserID <> "" Then
    Dim o_PM As New frmPrivMsg
    o_PM.Init Me, sToUserID
    o_PrivMsg.Add o_PM, sToUserID
    Set o_PM = Nothing
  Else
    GoTo AddPrivMsg_Resume
  End If
  
  AddPrivMsg = True
AddPrivMsg_Resume:
End Function


Public Function RemovePrivMsg(IndexKey) As Boolean
  On Error Resume Next
  o_PrivMsg.Remove IndexKey
  If Err = 0 Then RemovePrivMsg = True
End Function


Public Function PrivMsgExists(ByVal sKey As String) As Boolean
  Dim oTmp As frmPrivMsg
  On Error Resume Next
  Set oTmp = o_PrivMsg(sKey)
  If Err = 0 Then PrivMsgExists = True
  Set oTmp = Nothing
End Function


Public Property Get UserID() As String
  UserID = m_UserID
End Property
'^'
Public Property Let UserID(ByVal sUserID As String)
  m_UserID = sUserID
End Property


Private Sub Form_Load()
  GetClientSettings
  tmrWSU.Enabled = True
  tmrClient.Enabled = True
End Sub


Private Sub GetClientSettings()
  Dim cReg As New clsRegistry
  Dim oMnu As Menu
  
  For Each oMnu In mnuCmd
    oMnu.Tag = oMnu.Caption
  Next
  
  cReg.ClassKey = HKEY_LOCAL_MACHINE
  cReg.SectionKey = c_RegAppKey
  cReg.ValueKey = "LastID"
  m_LastID = Trim(cReg.Value)
  If Trim(m_LastID) = "" Then m_LastID = "CodeXP"
End Sub


Private Sub SaveClientSettings()
  Dim cReg As New clsRegistry
  
  cReg.ClassKey = HKEY_LOCAL_MACHINE
  cReg.SectionKey = c_RegAppKey
  cReg.ValueKey = "LastID"
  cReg.ValueType = REG_SZ
  cReg.Value = m_LastID
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If UnloadMode = 0 Then
'    Cancel = True
'    Me.Hide
'  End If
End Sub


Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  picSend.Move 0, MinScaleHeight - picSend.Height, MinScaleWidth
  txtSend.Move lblCMDL.Width, 0, picSend.ScaleWidth
  lstOut.Move 0, 0, MinScaleWidthUL, picSend.Top
  picUserList.Move lstOut.Width, 0, picUserList.Width, MinScaleHeight - picSend.Height
  lblWait.Move (ScaleWidth - lblWait.Width) / 2, picSend.Top + (picSend.Height - lblWait.Height) / 2
  If picUserList.Visible Then
    lstUsers.Move 0, 0, picUserList.ScaleWidth, picUserList.ScaleHeight
  End If
End Sub


Private Property Get MinScaleWidthUL() As Single
  Dim sW As Single
  sW = ScaleWidth - IIf(picUserList.Visible, picUserList.Width, 0)
  MinScaleWidthUL = IIf(sW > 1000, sW, 1000)
End Property


Private Property Get MinScaleWidth() As Single
  MinScaleWidth = IIf(ScaleWidth > 1000, ScaleWidth, 1000)
End Property


Private Property Get MinScaleHeight() As Single
  MinScaleHeight = IIf(ScaleHeight > 1000, ScaleHeight, 1000)
End Property


Private Sub Form_Unload(Cancel As Integer)
  Dim fPM As frmPrivMsg
  On Error Resume Next
  For Each fPM In o_PrivMsg
    Unload fPM
  Next fPM
  WSU.Close
  If Not f_Modules Is Nothing Then
    Unload f_Modules
    Set f_Modules = Nothing
  End If
  tmrClient.Enabled = False
  tmrWSU.Enabled = False
  DoMoreEvents
  tmrWSU_Timer
  DoMoreEvents
  bConnect = False
  bLoaded = False
  SaveClientSettings
End Sub


Private Sub lblCMDL_Click()
  Dim oMnu As Menu
  For Each oMnu In mnuCmd
    oMnu.Caption = Replace(oMnu.Tag, "%UserID%", m_LastID)
  Next
  PopupMenu menuCommands, 0, picSend.Left, picSend.Top
End Sub


Private Sub lstOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu menuMain
  End If
End Sub


Private Sub lstUsers_DblClick()
  Dim Tmp As String
  On Error Resume Next
  Tmp = lstUsers
  If Len(Tmp) > 0 Then
    If Not PrivMsgExists(Tmp) Then AddPrivMsg Tmp
    PrivMsg(Tmp).Show
  End If
End Sub


Private Sub mnuClose_Click()
  Unload Me
End Sub


Private Sub mnuCmd_Click(Index As Integer)
  On Error Resume Next
  txtSend = mnuCmd(Index).Caption
  txtSend.SelStart = Len(txtSend)
  txtSend.SetFocus
End Sub


Private Sub mnuModules_Click()
  If f_Modules Is Nothing Then
    Set f_Modules = New frmModules
    f_Modules.Init Me
  End If
  f_Modules.Show
End Sub


Private Sub tmrClient_Timer()
  Static bDone  As Boolean
  Static lWait  As Long
  Dim bEna As Boolean
  
  If Not bDone Then
    If Me.Visible Then
      If lWait > 5 Then
        On Error Resume Next
        txtSend.SetFocus
        If Err Then Err.Clear
        bDone = True
      End If
      lWait = lWait + 1
    End If
  End If
  
  bEna = (WSU.State = 7)
  If picSend.Visible <> bEna Then picSend.Visible = bEna
  If lblWait.Visible <> Not bEna Then lblWait.Visible = Not bEna
  If mnuModules.Enabled <> bEna Then mnuModules.Enabled = bEna
  
  bEna = (lstUsers.ListCount > 0)
  If picUserList.Visible <> bEna Then
    picUserList.Visible = bEna
    Form_Resize
  End If
End Sub


Public Sub ResetClient()
  m_UserID = ""
End Sub


Public Sub ExecCmd(ByVal CmdLine As String)
  Dim Tmp As String
  Dim oP  As New clsCXParser
  Dim i   As Long
  
  If ConnTime = 0 Then tmrWSU_Timer
  
  oP.Parse CmdLine
  
  Select Case UCase(oP.Class)
  Case "", "{ALL}"
    
    Select Case UCase(oP.Cmd)
    Case "MODULES"
      
      Tmp = ""
      If oP.ParamCount > 0 Then
        Select Case UCase(oP.Param(1))
        Case "LOAD"
          
          If oP.ParamCount > 1 Then
            If Not f_Modules Is Nothing Then f_Modules.lstMods.AddItem oP.Param(2)
          End If
        
        Case "UNLOAD"
          
          If oP.ParamCount > 1 Then
            If Not f_Modules Is Nothing Then
              For i = 0 To f_Modules.lstMods.ListCount - 1
                If UCase(f_Modules.lstMods.List(i)) = UCase(oP.Param(2)) Then
                  f_Modules.lstMods.RemoveItem i
                End If
              Next i
            End If
          End If
        
        Case "LIST"
        
          Tmp = "DONE"
          
        End Select
      Else
        Tmp = "DONE"
      End If
      
      If Tmp = "DONE" Then
        Dim cTok As New clsTokens
        Tmp = ""
        cTok.Init oP.Message, ";"
        
        If Not f_Modules Is Nothing Then f_Modules.lstMods.Clear
        For i = 1 To cTok.Count
          If Not f_Modules Is Nothing Then f_Modules.lstMods.AddItem cTok.Token(i)
        Next i
      End If
      
    Case "VERSION"
      
      If oP.ParamCount > 0 Then
        ' GOT VERSION OF OTHER SIDE '
      Else
        SendLine "VERSION " & myVersion
      End If
      
    Case "HELP"
      
      ' RECIEVE HELP HERE '
    
    Case "ERROR"
      
      ' RECIEVE ERRORS HERE '
      
    End Select
    
  Case "COMM"
  
    Select Case UCase(oP.Cmd)
    Case "PRIVMSG"
    
      Dim sFrom As String
      Dim sTo   As String
      Dim sPM   As String
      
      sFrom = oP.UserID
      If oP.ParamCount > 0 Then sTo = oP.Param(1)
      If (sFrom <> "") And (sTo <> "") Then
        
        If UCase(sFrom) = UCase(UserID) Then
          sPM = sTo
        Else
          sPM = sFrom
        End If
        
        If Not PrivMsgExists(sPM) Then AddPrivMsg sPM
        
        If Not PrivMsg(sPM).Visible Then
          On Error Resume Next
          PrivMsg(sPM).Show
        End If
        
        PrivMsg(sPM).AddMsg sFrom, oP.Message, oP.Params
      
      End If
      
    Case "SRVMSG"
      
      If oP.ParamCount > 0 Then
        Select Case UCase(oP.Param(1))
        Case "INFO"
          
          If oP.ParamCount > 1 Then
            Select Case UCase(oP.Param(2))
            Case "BEGIN"
              
              m_SrvInfo = oP.Message
            
            Case "END"
              
              If Len(m_SrvInfo) Then m_SrvInfo = m_SrvInfo & vbCrLf
              m_SrvInfo = m_SrvInfo & oP.Message
              
              Dim fNotice As New frmNotice
              fNotice.Caption = " Server Informations"
              fNotice.lblNotice = m_SrvInfo
              fNotice.Show
            
            End Select
          Else
            If Len(m_SrvInfo) Then m_SrvInfo = m_SrvInfo & vbCrLf
            m_SrvInfo = m_SrvInfo & oP.Message
          End If
        
        End Select
      Else
        ' SERVER MESSAGE RECEIVED '
      End If
    
    Case "LOGIN"
    
      If Len(oP.UserID) > 0 Then
        If oP.ParamCount > 0 Then
          Select Case UCase(oP.Param(1))
          Case "OK"
          
            UserID = oP.UserID
            m_LastID = UserID
            
          End Select
        End If
        AddUser oP.UserID
      End If
    
    Case "LEAVE"
    
      RemoveUser oP.UserID
    
    End Select
    
  End Select
  
  AddMsg ">>> " & CmdLine
End Sub


Public Sub AddUser(ByVal sUserID As String)
  Dim i As Long
  sUserID = Trim(sUserID)
  i = GetUserIndex(sUserID)
  If i < 0 Then
    lstUsers.AddItem sUserID
  End If
End Sub


Public Sub RemoveUser(ByVal sUserID As String)
  Dim i As Long
  sUserID = Trim(sUserID)
  i = GetUserIndex(sUserID)
  If i >= 0 Then
    lstUsers.RemoveItem i
  End If
End Sub


Public Function GetUserIndex(ByVal sUserID As String) As Long
  Dim i As Long
  GetUserIndex = -1
  sUserID = Trim(sUserID)
  For i = 0 To lstUsers.ListCount - 1
    If UCase(sUserID) = UCase(lstUsers.List(i)) Then
      GetUserIndex = i
      Exit For
    End If
  Next i
End Function


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
    AddMsg "<<< " & CmdLine
    WSU.SendData CmdLine & vbCrLf
    DoMoreEvents  ' System Events '
  Else
    ' We are not Connected: Log Error '
    LogError "Error: Can't send because not connected"
    Stop
  End If
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
      'LogEvent Format(m_Index, "\[000\] ") & WSU.RemoteHostIP & " Connected"
      SendLine "Version"
      SendLine "Modules"
    End If
    
  ' Socket is Closed                                                      '
  Case 0
  
    ' Check if we was connected  '
    If ConnTime <> 0 Then
      If Not f_Modules Is Nothing Then
        Unload f_Modules
        Set f_Modules = Nothing
      End If
      ' Log disconnect Event                     '
      'LogEvent Format(m_Index, "\[000\] ") & WSU.RemoteHostIP & " Closed"
      WSU_DataArrival -1  ' Clear Buffer  '
      ConnTime = 0
      lstUsers.Clear
    End If
    
    ' If we try to Connect to Server  '
    If bConnect Then
      bConnect = False
      ' Error occured on connecting  '
      lblWait = "Connection failed!"
      Form_Resize
      'LngMsgBox 100
    End If
  
  ' All Other States                                                      '
  Case Else
  
    ' Close Socket  '
    WSU.Close
    
  End Select
End Sub


Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim Tmp As String
  Select Case KeyCode
  Case vbKeyReturn
    Tmp = txtSend
    txtSend = ""
    SendLine Tmp
    KeyCode = 0
  End Select
End Sub


Private Sub WSU_DataArrival(ByVal bytesTotal As Long)
  Static Buffer As String
  Static lInstances As Long
  Dim oT  As New clsTokens
  Dim Tmp As String
  Dim i   As Long
  
  If bytesTotal < 0 Then
    ' Reset Buffer  '
    Buffer = ""
    Exit Sub
  End If
  
  ' Count up recursive Instances  '
  lInstances = lInstances + 1
  While lInstances > 100
    ' Wait here until Instances < 101 '
    DoEvents
  Wend
  
  ' Count up recieved Data Bytes  '
  m_DataIn = m_DataIn + bytesTotal
  
  ' Get recieved Data '
  WSU.GetData Tmp
  If Len(Buffer) Then
    ' Prepend Buffer on Data  '
    Tmp = Buffer & Tmp
    Buffer = ""
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
End Sub


Private Sub WSU_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  ' Log Error '
  LogError "Socket Error: " & vbCrLf & Description
End Sub


Public Function SocketFree() As Boolean
  SocketFree = (ConnTime = 0) And (WSU.State = 0)
End Function


Public Function Connect(ByVal sHOST As String, Optional ByVal lPort As Long = 105) As Boolean
  On Error Resume Next
  If WSU.State Then WSU.Close
  tmrWSU_Timer
  DoMoreEvents
  ' Try to connect to Host      '
  WSU.Connect sHOST, lPort
  bConnect = True
  If Err = 0 Then Connect = True
End Function
