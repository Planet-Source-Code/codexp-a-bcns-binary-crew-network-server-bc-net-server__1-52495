VERSION 5.00
Begin VB.Form frmPrivMsg 
   AutoRedraw      =   -1  'True
   Caption         =   "Private Message Window"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "cliPrivMsg.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   2415
   ScaleWidth      =   4695
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   735
      Begin VB.Timer tmrPrivMsg 
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
      TabIndex        =   3
      Top             =   0
      Width           =   4575
   End
   Begin VB.PictureBox picSend 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4545
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtSend 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   1
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
End
Attribute VB_Name = "frmPrivMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_UserID  As String
Private f_Client  As frmClient



Public Property Get UserID() As String
  UserID = m_UserID
End Property
'^'
Public Property Let UserID(ByVal sUserID As String)
  m_UserID = sUserID
End Property


Public Property Get FromUserID() As String
  FromUserID = f_Client.UserID
End Property


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Cancel = True
    Me.Hide
  End If
End Sub


Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  picSend.Move 0, MinScaleHeight - picSend.Height, MinScaleWidth
  txtSend.Move lblCMDL.Width, 0, picSend.ScaleWidth
  lstOut.Move 0, 0, MinScaleWidth, picSend.Top
  lblWait.Move (ScaleWidth - lblWait.Width) / 2, picSend.Top + (picSend.Height - lblWait.Height) / 2
End Sub


Private Property Get MinScaleWidth() As Single
  MinScaleWidth = IIf(ScaleWidth > 1000, ScaleWidth, 1000)
End Property


Private Property Get MinScaleHeight() As Single
  MinScaleHeight = IIf(ScaleHeight > 1000, ScaleHeight, 1000)
End Property


Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  f_Client.RemovePrivMsg UserID
End Sub


Public Sub AddLine(ByVal sMsg As String)
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


Public Sub AddMsg(ByVal sFrom As String, ByVal sMsg As String, Optional ByVal sParams As String)
  AddLine "<" & sFrom & "> " & sMsg
End Sub


Public Sub SendMsg(ByVal sToUser As String, ByVal sMsg As String, Optional ByVal sParams As String)
  f_Client.SendLine ":" & UserID & " Comm.PrivMsg " & sToUser & " " & Trim(sParams) & ":" & sMsg
End Sub


Private Sub tmrPrivMsg_Timer()
  Dim bEna As Boolean
  bEna = (f_Client.WSU.State = 7)
  If picSend.Visible <> bEna Then picSend.Visible = bEna
End Sub


Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim Tmp As String
  Select Case KeyCode
  Case vbKeyReturn
    Tmp = txtSend
    txtSend = ""
    SendMsg UserID, Tmp
    KeyCode = 0
  End Select
End Sub


Public Sub Init(fClient As frmClient, ByVal sToUser As String)
  Set f_Client = fClient
  UserID = sToUser
  tmrPrivMsg.Enabled = True
  Me.Caption = " [" & UserID & "] < " & FromUserID
End Sub

