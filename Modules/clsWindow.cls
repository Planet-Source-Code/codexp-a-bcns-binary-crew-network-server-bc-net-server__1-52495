VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const GWL_EXSTYLE = (-20)   ' Sets a new extended window style.           '
Private Const GWL_STYLE = (-16)     ' Sets a new window style.                    '
Private Const GWL_WNDPROC = (-4)    ' Sets a new address for the window procedure.'
Private Const GWL_HINSTANCE = (-6)  ' Sets a new application instance handle.     '
Private Const GWL_ID = (-12)        ' Sets a new identifier of the window.        '
Private Const GWL_USERDATA = (-21)  ' Sets the 32-bit value associated with window'

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private m_hWnd    As Long
Private m_Parent  As clsWindow

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long


' hWnd Property                                         (C)2003 by CodeXP '
Public Property Get hWnd() As Long
  hWnd = m_hWnd
End Property
Public Property Let hWnd(ByVal SetToHWnd As Long)
  m_hWnd = SetToHWnd
End Property


' ClassName Property                                    (C)2003 by CodeXP '
Public Property Get ClassName() As String
  ClassName = myClassName
End Property


' Caption Property                                      (C)2003 by CodeXP '
Public Property Get Caption() As String
  Caption = GetWindowCaption(m_hWnd)
End Property
Public Property Let Caption(ByVal NewCaption As String)
  SetWindowCaption m_hWnd, NewCaption
End Property


' Parent Property                                       (C)2003 by CodeXP '
Public Property Get Parent() As clsWindow
  If m_Parent Is Nothing Then
    Set m_Parent = New clsWindow
  End If
  m_Parent.hWnd = GetParent(m_hWnd)
  Set Parent = m_Parent
End Property
Public Property Let Parent(ByVal NewParent As Variant)
  If IsObject(NewParent) Then
    Set m_Parent = NewParent
    SetParent m_hWnd, NewParent.hWnd
  Else
    SetParent m_hWnd, NewParent
  End If
End Property


' ClientWidth Property                                  (C)2003 by CodeXP '
Public Property Get ClientWidth()
  Dim WRect As RECT
  GetWindowRect m_hWnd, WRect
  ClientWidth = WRect.Right - WRect.Left
End Property


' ClientHeight Property                                 (C)2003 by CodeXP '
Public Property Get ClientHeight()
  Dim WRect As RECT
  GetWindowRect m_hWnd, WRect
  ClientHeight = WRect.Bottom - WRect.Top
End Property


' FindClass() - Sucht die Klasse und gibt hWnd zurück   (C)2003 by CodeXP '
Public Function FindClass(ByVal ClassName As String) As Long
  FindClass = FindWindow(ClassName, vbNullString)
End Function










' GetWindowCaption() - Liest Fenstertext aus            (C)2003 by CodeXP '
Private Function GetWindowCaption(ByVal hWnd As Long) As String
  Dim sBuffer As String
  Dim lsLen As Long
  lsLen = 255
  sBuffer = String(lsLen, " ")
  lsLen = GetWindowText(hWnd, sBuffer, lsLen)
  GetWindowCaption = Left(sBuffer, lsLen)
End Function


' SetWindowCaption() - Setzt neuen Fenstertext          (C)2003 by CodeXP '
Private Function SetWindowCaption(ByVal hWnd As Long, ByVal Cap As String) As Long
  SetWindowCaption = SetWindowText(hWnd, Cap & Chr(0))
End Function


' myClassName() - Liest Classenname aus                 (C)2003 by CodeXP '
Private Function myClassName() As String
  Dim sBuffer As String
  Dim lsLen As Long
  lsLen = 255
  sBuffer = String(lsLen, " ")
  lsLen = GetClassName(m_hWnd, sBuffer, lsLen)
  myClassName = Left(sBuffer, lsLen)
End Function
