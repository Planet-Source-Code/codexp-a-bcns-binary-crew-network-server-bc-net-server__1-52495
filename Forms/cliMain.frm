VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "BCC Connect"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "cliMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   2760
   Begin VB.Frame Frame1 
      Caption         =   "[ Connect ]"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect ..."
         Default         =   -1  'True
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "105"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConnect_Click()
  Dim i As Long
  Dim l As Long
  
  On Error Resume Next
  
  For i = 0 To c_MaxClients - 1
    If o_Cleints(i) Is Nothing Then Set o_Cleints(i) = New frmClient
    If Not o_Cleints(i).bLoaded Then
      l = i + 1
      Exit For
    End If
  Next i
  
  If l > 0 Then
    i = l - 1
    Load o_Cleints(i)
    o_Cleints(i).bLoaded = True
    o_Cleints(i).Show
    If Not o_Cleints(i).Connect(txtServer, Val(txtPort)) Then
      o_Cleints(i).bLoaded = False
      Unload o_Cleints(i)
      LngMsgBox 100
    End If
  End If
End Sub


Private Sub Form_Load()
  Dim TrayWnd As New clsWindow
  Dim MyWnd   As New clsWindow
  
  Me.Show
  
  MyWnd.hWnd = Me.hWnd
  TrayWnd.hWnd = TrayWnd.FindClass("Shell_TrayWnd")
  Me.Left = Screen.Width - (MyWnd.ClientWidth) * Screen.TwipsPerPixelX
  Me.Top = Screen.Height - (MyWnd.ClientHeight + TrayWnd.ClientHeight) * Screen.TwipsPerPixelY
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
