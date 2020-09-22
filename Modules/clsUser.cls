VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public UserID As String
Public Index  As Long

Private SAI   As clsNetAppInterface


Public Sub Init(SAIRef As clsNetAppInterface, ByVal lIndex As Long, Optional ByVal sUserID As String)
  Index = lIndex
  Set SAI = SAIRef
  If Len(sUserID) Then UserID = sUserID
End Sub


Public Sub SendLine(ByVal CmdLine As String)
  On Error Resume Next
  SAI.SendLine Index, CmdLine
End Sub
