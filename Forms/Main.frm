VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "BC Net Server"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2625
      ScaleWidth      =   2745
      TabIndex        =   3
      Top             =   120
      Width           =   2775
      Begin VB.Frame fraConnections 
         BackColor       =   &H00FFFFFF&
         Caption         =   "[          -               ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   2535
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   " Connections "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   21
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label lblConnCount 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   20
            Top             =   480
            Width           =   90
         End
         Begin VB.Label lblFree 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   19
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Free Sockets:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2250
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Connections:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   2250
         End
      End
      Begin VB.Frame fraInformation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "[                          ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   2535
         Begin VB.Label lblLoadedModules 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   22
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   " Information "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   15
            Top             =   0
            Width           =   1110
         End
         Begin VB.Label lblWSState 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   14
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Server Status:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   2250
         End
         Begin VB.Label lblSrvUpTime 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   12
            Top             =   480
            Width           =   90
         End
         Begin VB.Label lblAppUpTime 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   11
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Application UpTime:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2250
         End
         Begin VB.Label lblDataIn 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   9
            Top             =   960
            Width           =   90
         End
         Begin VB.Label lblDataOut 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   8
            Top             =   1200
            Width           =   90
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Server UpTime:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   2250
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Data In:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   2250
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Total Data Out:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   2250
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Loaded Modules:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   2250
         End
      End
   End
   Begin VB.CommandButton cmdServer 
      Caption         =   "Stop Server"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect ..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   1665
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Timer tmrApp 
         Enabled         =   0   'False
         Interval        =   333
         Left            =   120
         Top             =   0
      End
      Begin VB.Timer tmrWS 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1080
         Top             =   0
      End
      Begin MSWinsockLib.Winsock WS 
         Left            =   600
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Label lblDblClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click on Form for Debug Wnd."
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   3360
      Width           =   2685
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConnect_Click()
  Dim Tmp As String
  Dim i   As Long
  Dim b   As Boolean
  
  Tmp = Trim(InputBox("IP oder Host eingeben:", "IP", WS.LocalIP))
  If Len(Tmp) > 0 Then
    If Not b Then LngMsgBox 1001
  End If
End Sub


Private Sub cmdServer_Click()
  If m_StartServer Then
    m_StartServer = False
    If WS.State = 2 Then WS.Close
  Else
    m_StartServer = True
    If WS.State Then WS.Close
  End If
End Sub


Private Sub Form_DblClick()
  frmDebug.Show
End Sub


Private Sub Form_Load()
  Dim i As Long
  
  Set o_Application = New clsApplication
  
  m_AppUpTime = TimeLong
  LogEvent "Application Start"
  
  For i = 0 To c_MaxUsers - 1
    Set o_Clients(i) = New frmClient
    o_Clients(i).Index = i
  Next i
  
  m_StartServer = True
  
  tmrApp.Enabled = True
  tmrWS.Enabled = True
  
  o_Modules.ReadModules
  o_Modules.StartModules
  
  LogEvent IIf(o_Modules.ModulesCount > 0, o_Modules.ModulesCount, "No") & " Modules Loaded"
  
  Me.Show
  Me.Top = 0
  Me.Left = Screen.Width - Width
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long
  
  CloseServer
    
  tmrApp.Enabled = False
  tmrWS.Enabled = False
  
  tmrWS_Timer
  
  For i = 0 To c_MaxUsers - 1
    Unload o_Clients(i)
  Next i
  
  LogEvent "Application Terminated"
  
  DoMoreEvents
  Unload frmDebug
  DoMoreEvents
  
  End
End Sub


Private Sub lblDblClick_DblClick()
  Form_DblClick
End Sub


Private Sub tmrApp_Timer()
  Dim Tmp As String
  Dim i   As Long
  Dim lCF As Long
  Dim lCC As Long
  
  For i = 0 To c_MaxUsers - 1
    Select Case o_Clients(i).WSU.State
    Case 7
      lCC = lCC + 1
    Case 0
      lCF = lCF + 1
    End Select
  Next i
  m_ClientCount = lCC
  m_FreeCount = lCF
  
  If m_AppUpTime > 0 Then
    Tmp = TimeLeftAfter(m_AppUpTime)
  Else
    Tmp = "0:00:00"
  End If
  If lblAppUpTime <> Tmp Then lblAppUpTime = Tmp
  
  If m_ServerUpTime > 0 Then
    Tmp = TimeLeftAfter(m_ServerUpTime)
  Else
    Tmp = "0:00:00"
  End If
  If lblSrvUpTime <> Tmp Then lblSrvUpTime = Tmp
  
  Tmp = BytesView(m_DataIn)
  If lblDataIn <> Tmp Then lblDataIn = Tmp
  
  Tmp = o_Modules.ModulesCount
  If lblLoadedModules <> Tmp Then lblLoadedModules = Tmp
  
  Tmp = BytesView(m_DataOut)
  If lblDataOut <> Tmp Then lblDataOut = Tmp
  
  Tmp = m_ClientCount
  If lblConnCount <> Tmp Then lblConnCount = Tmp
  
  Tmp = m_FreeCount & "/" & c_MaxUsers
  If lblFree <> Tmp Then lblFree = Tmp
  
  Tmp = IIf(WS.State = 2, "Stop Server", "Start Server")
  If cmdServer.Caption <> Tmp Then cmdServer.Caption = Tmp
  
  Tmp = GetSocketState(WS.State)
  If lblWSState <> Tmp Then lblWSState = Tmp
End Sub


Private Sub tmrWS_Timer()
  Static OldState As Long
  Dim CurWSState  As Long
  
  CurWSState = WS.State
  
  Select Case CurWSState
  Case 2    ' Listeninig      '
        
    If OldState <> 2 Then
      m_ServerUpTime = TimeLong
      LogEvent "Server Started on Port " & WS.LocalPort
    End If
          
  Case 0  ' Idle  '
    
    If OldState <> 0 Then
      m_ServerUpTime = 0
      LogEvent "Server Socket Closed"
    End If
    
    If m_StartServer Then StartServer
  
  Case Else
  
    WS.Close
  
  End Select
  
  OldState = CurWSState
End Sub


Private Sub StartServer()
  On Error Resume Next
  WS.Close
  WS.LocalPort = c_NetPort
  WS.Listen
  If Err Then
    m_StartServer = False
    LogError "Server can not be Started"
    Err.Clear
  End If
End Sub


Private Sub CloseServer()
  WS.Close
  WS.LocalPort = 0
End Sub


Private Sub WS_ConnectionRequest(ByVal requestID As Long)
  Dim i As Long
  Dim b As Boolean
  
  If GetPFaults(WS.RemoteHostIP) > 10 Then
    Exit Sub
  End If
  
  For i = 0 To c_MaxUsers - 1
    If o_Clients(i).SocketFree Then
      o_Clients(i).bConnect = False
      o_Clients(i).WSU.Accept requestID
      b = True
      Exit For
    End If
  Next i
  
  If Not b Then LogError "Connection Refused for " & WS.RemoteHostIP & " (Server Full)"
End Sub


Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  LogError "Server Socket Error: " & vbCrLf & Description
End Sub



