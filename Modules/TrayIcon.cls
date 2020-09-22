VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' TrayIcon.cls (C)2003 by CodeXP                                          '
Option Explicit

' Constants                                     '
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const MAX_Tip_LENGTH = 64&

' Types                                         '
Private Type NOTIFYICONDATA
  cbSize            As Long
  hWnd              As Long
  uID               As Long
  uFlags            As Long
  uCallbackMessage  As Long
  hIcon             As Long
  szTip             As String * MAX_Tip_LENGTH
End Type

' Events                                        '
Public Event Error(ByVal ErrorNumber As Long)
Public Event DblClick(ByVal Button As Long)
Public Event MouseDown(ByVal Button As Long)
Public Event MouseUp(ByVal Button As Long)
Public Event MouseMove()

' Variables                                     '
Private STIN As NOTIFYICONDATA
Private bDisplayIcon As Boolean
Private bUpdateChange As Boolean

' API's                                         '
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


Public Function Init(ByVal hWnd As Long, ByVal hIcon As Long, ByVal sTip As String, Optional ByVal uCallbackMsg As Long = WM_MOUSEMOVE) As Long
  With STIN
    .hWnd = hWnd
    .uID = vbNull
    .hIcon = hIcon
    .cbSize = Len(STIN)
    .uCallbackMessage = uCallbackMsg
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .szTip = Left(sTip, MAX_Tip_LENGTH - 1) & vbNullChar
  End With
  bDisplayIcon = False
  bUpdateChange = True
End Function


Public Function Show() As Boolean
  If Not bDisplayIcon Then
    Show = Shell_NotifyIcon(NIM_ADD, STIN)
    If Show = False Then
      RaiseEvent Error(GetLastError)
    Else
      bDisplayIcon = True
    End If
  End If
End Function


Public Function Hide() As Boolean
  If bDisplayIcon Then
    Hide = Shell_NotifyIcon(NIM_DELETE, STIN)
    If Hide = False Then
      RaiseEvent Error(GetLastError)
    Else
      bDisplayIcon = False
    End If
  End If
End Function


Public Function Update() As Boolean
  If bDisplayIcon Then
    Update = Shell_NotifyIcon(NIM_MODIFY, STIN)
    If Update = False Then
      RaiseEvent Error(GetLastError)
    End If
  End If
End Function


Public Function RaiseEventFromMsg(ByVal lMsg As Long, ByVal lParam As Long) As Long
  RaiseEventFromMsg = lMsg
  Select Case lMsg
  Case WM_LBUTTONDBLCLK
    RaiseEvent DblClick(1)
  Case WM_LBUTTONDOWN
    RaiseEvent MouseDown(1)
  Case WM_LBUTTONUP
    RaiseEvent MouseUp(1)
  Case WM_MBUTTONDBLCLK
    RaiseEvent DblClick(4)
  Case WM_MBUTTONDOWN
    RaiseEvent MouseDown(4)
  Case WM_MBUTTONUP
    RaiseEvent MouseUp(4)
  Case WM_MOUSEMOVE
    RaiseEvent MouseMove
  Case WM_RBUTTONDBLCLK
    RaiseEvent DblClick(2)
  Case WM_RBUTTONDOWN
    RaiseEvent MouseDown(2)
  Case WM_RBUTTONUP
    RaiseEvent MouseUp(2)
  Case Else
    RaiseEventFromMsg = 0
  End Select
End Function


Public Property Get IconHandle() As Long
  IconHandle = STIN.hIcon
End Property
Public Property Let IconHandle(ByVal hIcon As Long)
  STIN.hIcon = hIcon
  If bUpdateChange Then
    STIN.uFlags = NIF_ICON
    Update
    STIN.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property


Public Property Get ToolTextTip() As String
  ToolTextTip = Left(STIN.szTip, Len(STIN.szTip) - 1)
End Property
Public Property Let ToolTextTip(ByVal sTip As String)
  STIN.szTip = Left(sTip, MAX_Tip_LENGTH - 1) & vbNullChar
  If bUpdateChange Then
    STIN.uFlags = NIF_TIP
    Update
    STIN.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property


Public Property Get CallBackMsg() As Long
  CallBackMsg = STIN.uCallbackMessage
End Property
Public Property Let CallBackMsg(ByVal uCallbackMsg As Long)
  STIN.uCallbackMessage = uCallbackMsg
  If bUpdateChange Then
    STIN.uFlags = NIF_MESSAGE
    Update
    STIN.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property


Public Property Let UpdateChange(bUpdate As Boolean)
  bUpdateChange = bUpdate
End Property
Private Property Get UpdateChange() As Boolean
  UpdateChange = bUpdateChange
End Property


Public Property Get IconVisible() As Boolean
  IconVisible = bDisplayIcon
End Property
Public Property Let IconVisible(ByVal bVisible As Boolean)
  If bVisible Then
    Show
  Else
    Hide
  End If
End Property


Private Sub Class_Terminate()
  Hide
End Sub

