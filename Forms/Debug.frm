VERSION 5.00
Begin VB.Form frmDebug 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   " Log Window"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Debug.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.ListBox lstOut 
      ForeColor       =   &H00FF0000&
      Height          =   2595
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Tag             =   "General"
      Top             =   0
      Width           =   4575
   End
   Begin VB.ListBox lstOut 
      Height          =   2595
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Tag             =   "Commands"
      Top             =   0
      Width           =   4575
   End
   Begin VB.ListBox lstOut 
      ForeColor       =   &H000000FF&
      Height          =   2595
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Tag             =   "Errors"
      Top             =   0
      Width           =   4575
   End
   Begin VB.Menu menuSwitch 
      Caption         =   "Switch"
      Visible         =   0   'False
      Begin VB.Menu mnuListBox 
         Caption         =   "General"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurListBox As String


Public Sub LogError(ByVal Msg As String)
  LogMsg Msg, "Errors"
End Sub


Public Sub LogCommand(ByVal Msg As String)
  LogMsg Msg, "Commands"
End Sub


Public Sub LogMessage(ByVal Msg As String)
  LogMsg Msg, "General"
End Sub


Private Sub LogMsg(ByVal Msg As String, ByVal ListBoxName As String)
  Dim oT  As New clsTokens
  Dim li  As Long
  Dim i   As Long
  
  li = FindListBox(ListBoxName)
  If li > 0 Then
    oT.Init Msg, vbCrLf, , True
    For i = 1 To oT.Count
      If i = oT.Count And oT.Token(i) = "" Then Exit For
      lstOut(li - 1).AddItem IIf(i = 1, TimeStamp, "") & oT.Token(i)
    Next i
    
    While lstOut(li - 1).ListCount > 1000
      lstOut(li - 1).RemoveItem 0
    Wend
    
    If lstOut(li - 1).ListCount > 0 Then
      lstOut(li - 1).TopIndex = lstOut(li - 1).ListCount - 1
    End If
  End If
End Sub


Private Sub Form_Load()
  Dim i As Long
  For i = 1 To lstOut.UBound
    Load mnuListBox(i)
    mnuListBox(i).Caption = lstOut(i).Tag
    mnuListBox(i).Visible = True
  Next i
  lstOut(0).ZOrder
  Me.Move frmMain.Left, frmMain.Top + frmMain.Height
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Cancel = True
    Me.Hide
  End If
End Sub


Private Function FindListBox(ByVal N As String) As Long
  Dim i As Long
  If Trim(N) = "" Then N = "General"
  For i = 0 To lstOut.UBound
    If UCase(Trim(lstOut(i).Tag)) = UCase(Trim(N)) Then
      FindListBox = i + 1
      Exit For
    End If
  Next i
End Function


Private Sub lstOut_DblClick(Index As Integer)
  Dim Msg As String
  Msg = lstOut(Index)
  If Len(Msg) > 50 Then
    MsgBox Msg, vbInformation, "Long Line View"
  End If
End Sub


Private Sub lstOut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  CurListBox = lstOut(Index).Tag
  If Button = vbRightButton Then
    PopupMenu menuSwitch
  End If
End Sub


Private Sub mnuListBox_Click(Index As Integer)
  lstOut(Index).ZOrder
End Sub
