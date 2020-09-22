VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private o_Module      As Object


Public Property Get Started() As Boolean
  On Error Resume Next
  Started = o_Module.Started
End Property


Public Function Init(oModuleRef As Object) As Boolean
  Set o_Module = oModuleRef
End Function


Public Function Execute(ByVal Index As Long, ByVal CmdLine As String) As Boolean
  On Error Resume Next
  Execute = o_Module.Execute(Index, CmdLine)
End Function


Public Property Get ModuleName() As String
  On Error Resume Next
  ModuleName = o_Module.ModuleName
End Property


Public Function StartModule() As Boolean
  On Error Resume Next
  Call o_Module.StartModule(o_Application)
  StartModule = Started
End Function


Public Sub Disconnected(ByVal Index As Long)
  On Error Resume Next
  Call o_Module.Disconnected(Index)
End Sub


Public Sub Connected(ByVal Index As Long)
  On Error Resume Next
  Call o_Module.Connected(Index)
End Sub


