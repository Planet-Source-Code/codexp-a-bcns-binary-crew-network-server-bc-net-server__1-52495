VERSION 5.00
Begin VB.Form frmModules 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Modules"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrModules 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start/&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load ..."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "&Unload"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox lstMods 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private f_Client    As frmClient


Private Sub cmdLoad_Click()
  Dim fInp As New frmInput
  Dim Tmp As String
  fInp.Show
  fInp.lblCaption = "Input Module Name:"
  While fInp.Visible
    DoMoreEvents
  Wend
  Tmp = Trim(fInp.sInput)
  Set fInp = Nothing
  If Tmp <> "" Then f_Client.SendLine "Modules Load " & Tmp
End Sub


Private Sub cmdUnload_Click()
  Dim Tmp As String
  Tmp = Trim(lstMods)
  If Len(Tmp) > 0 Then
    f_Client.SendLine "Modules Unload " & Tmp
  End If
End Sub


Public Sub Init(fClient As frmClient)
  Set f_Client = fClient
  tmrModules.Enabled = True
  f_Client.SendLine "Modules List"
End Sub


Private Sub Form_Unload(Cancel As Integer)
  tmrModules.Enabled = False
  DoMoreEvents
End Sub


Private Sub tmrModules_Timer()
  Dim Tmp As String
  Tmp = "&Start/&Stop"
  If cmdStart.Caption <> Tmp Then cmdStart.Caption = Tmp
End Sub
