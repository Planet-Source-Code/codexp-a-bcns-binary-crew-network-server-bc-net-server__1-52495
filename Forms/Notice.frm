VERSION 5.00
Begin VB.Form frmNotice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   " Notice"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.PictureBox picNotice 
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.Label lblNotice 
         BackStyle       =   0  'Transparent
         Caption         =   "Notice"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  picNotice.Move picNotice.Left, picNotice.Top, MinScaleWidth - picNotice.Left * 2, MinScaleHeight - picNotice.Top * 2
  lblNotice.Move 0, 0, picNotice.ScaleWidth, picNotice.ScaleHeight
End Sub


Private Property Get MinScaleWidth() As Single
  MinScaleWidth = IIf(ScaleWidth > 1000, ScaleWidth, 1000)
End Property


Private Property Get MinScaleHeight() As Single
  MinScaleHeight = IIf(ScaleHeight > 1000, ScaleHeight, 1000)
End Property

