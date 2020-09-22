VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Input"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2160
   StartUpPosition =   1  'Fenstermitte
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Input:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   405
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sInput As String


Private Sub cmdCancel_Click()
  sInput = ""
  Me.Hide
End Sub


Private Sub cmdOk_Click()
  sInput = txtInput
  Me.Hide
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Cancel = True
    cmdCancel_Click
  End If
End Sub
