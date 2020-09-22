VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const c_ModuleName = "SDK"

Private SAI         As New clsNetAppInterface ' Server Application Interface  '
Private m_bStarted  As Boolean                ' Module Started (True/False)   '


' Interface Property: Started()                                                         '
Public Property Get Started() As Boolean
  ' Return the State of Module '
  Started = m_bStarted
End Property


' Interface Property: ModuleName()                                                      '
Public Property Get ModuleName() As String
  ' Return the Name of this Module '
  ModuleName = c_ModuleName
End Property


' Interface Property: ModuleVersion()                                                   '
Public Function ModuleVersion() As String
  ' Return Version of this Module '
  ModuleVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function


' Interface Event: Connected()                                                          '
Public Sub Connected(ByVal Index As Long)
  ' Event raises on connecting Socket(Index) '
End Sub


' Interface Event: Disconnected()                                                       '
Public Sub Disconnected(ByVal Index As Long)
  ' Event raises on closing Socket(Index) '
End Sub


' Interface Action: StartModule()                                                       '
Public Function StartModule(oNetApp As Object) As Boolean
  ' Here we have to Start our Module.   '
  '> Return True if Module was started, '
  '  else Return False if not.          '
  '> We have to set a Reference to the  '
  '  Server Application Interface       '
  '  (oNetApp).                         '
  
  On Error GoTo StartModule_Resume
  
  ' If already started then Exit: '
  If m_bStarted Then
    ' Return True: '
    StartModule = True
    Exit Function
  End If
  
  ' Set Reference to the Server Application Interface '
  If Not SAI.Init(oNetApp) Then Exit Function
  
  ' Check Server Version  '
  If Val(SAI.Version) < Val(ModuleVersion) Then
    ' Wrong Server Version! '
  Else
    ' Return True: '
    StartModule = True
  End If

StartModule_Resume:
  m_bStarted = StartModule
End Function


' Interface Action: StopModule()                                                        '
Public Function StopModule(oNetApp As Object) As Boolean
  ' Here we have to Stop our Module.    '
End Function


' Interface Event: Execute()                                                            '
Public Function Execute(ByVal Index As Long, ByVal CmdLine As String) As Boolean
  ' Event raises on DataArrival on Socket(Index)    '
  '> Here we have to Parse and Execute the          '
  '  Command Line (CmdLine).                        '
  '> We can send an answer with SendLine() Function.'
  '> Return False if Command is unknown,            '
  '  else Return True if Command is known.          '
  
  Dim oP  As New clsCXParser
  Dim Tmp As String
  Dim i   As Long
  
  ' Parse Command Line: '
  oP.Parse CmdLine
  
  ' If Module is not started then Exit: '
  If Not m_bStarted Then Exit Function
  ' If Class does not match ModuleName then Exit: '
  If (UCase(oP.Class) <> UCase(ModuleName)) Then Exit Function
  
  Execute = True
  ' Execute Command: '
  Select Case UCase(oP.Cmd)
  '>Command: "Version"      '
  Case "VERSION"
    
    ' If any Params exists then: '
    If oP.ParamCount > 1 Then
      
      ' Version String recieved from Client '
      Debug.Print "Client Version: " & oP.Param(1)
    
    Else ' if no Params then: '
      
      ' Send Version String: '
      SAI.SendLine Index, ModuleName & ".Version " & ModuleVersion & _
                          ":Interface Module Copyright (C)2004 by CodeXP"
    
    End If
    
  '>Unknown Command         '
  Case Else
  
    ' Return False: '
    Execute = False
    
  End Select
  
End Function

