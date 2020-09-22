VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsModule"
Attribute VB_Ext_KEY = "Member0" ,"clsModule"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection


Public Function ReadModules(Optional ByVal OnlyModule As String) As Integer
  Dim lModAttr  As VbFileAttribute
  Dim sModPatt  As String
  Dim sModPath  As String
  Dim sCurFile  As String
  Dim oModule   As Object
  Dim bFail     As Boolean
  
  On Error Resume Next
  
  ' Set File Pattern for Modules    '
  sModPatt = "*.DLL"
  ' Set Path for Modules            '
  sModPath = myPath & "Extensions\"
  ' Set File Attribute for Modules  '
  lModAttr = vbArchive Or vbHidden Or vbReadOnly Or vbSystem
  
  ' Start Searching Module Files    '
  bFail = True
  sCurFile = Dir(sModPath & sModPatt, lModAttr)
  While Len(sCurFile)
    ' Load Module                   '
    If Not (LoadModule(sModPath, sCurFile, OnlyModule) Is Nothing) Then
      bFail = False
    End If
    ' Get next Module File          '
    sCurFile = Dir
  Wend
  
  If Not bFail Then ReadModules = mCol.Count
End Function


Public Function LoadModule(ByVal sModPath As String, ByVal sModFile As String, Optional ByVal OnlyModule As String) As Object
  Dim sModName  As String
  
  Shell "RegSvr32.exe /S " & sModPath & sModFile, vbHide
  If Err Then Err.Clear
  DoMoreEvents
  
  Set LoadModule = CreateObject(ModuleName(sModFile) & ".Module")
  sModName = LoadModule.ModuleName
  If Len(OnlyModule) > 0 Then
    If UCase(sModName) <> UCase(OnlyModule) Then
      Set LoadModule = Nothing
      Exit Function
    End If
  End If
  
  If Err Then
    Set LoadModule = Nothing
    sModName = ""
    Err.Clear
  Else
    If Not AddModule(LoadModule) Is Nothing Then
      LogEvent "Module """ & sModName & """ loaded"
    End If
  End If
End Function


Public Sub StartModules()
  Dim oMod As clsModule
  For Each oMod In mCol
    If Not oMod.StartModule Then
      RemoveModule oMod.ModuleName
    End If
  Next oMod
End Sub


Private Function AddModule(oModuleRef As Object) As clsModule
  Dim oNewMod As clsModule
  
  On Error Resume Next
  
  Set oNewMod = New clsModule
  Call oNewMod.Init(oModuleRef)
  
  If Err Then
    Set oNewMod = Nothing
    Err.Clear
  Else
    mCol.Add oNewMod, oNewMod.ModuleName()
  End If

  Set AddModule = oNewMod
  Set oNewMod = Nothing
End Function


Public Property Get Module(IndexKey) As clsModule
Attribute Module.VB_UserMemId = 0
  Set Module = mCol(IndexKey)
End Property



Public Property Get ModulesCount() As Long
  ModulesCount = mCol.Count
End Property


Public Sub RemoveModule(IndexKey)
  mCol.Remove IndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  ' Diese Eigenschaft ermöglicht, die Elemente der  '
  ' Auflistung mit For...Each zu durchlaufen.       '
  Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub


Public Function ModuleExists(IndexKey) As Boolean
  Dim oMod As clsModule
  
  On Error Resume Next
  Set oMod = mCol(IndexKey)
  If Err Then Exit Function
  
  ModuleExists = True
End Function


' Interface Raise Event: Disconnected()                                                 '
Public Sub Disconnected(ByVal Index As Long)
  Dim oMod As clsModule
  On Error Resume Next
  For Each oMod In mCol
    oMod.Disconnected Index
  Next oMod
End Sub


' Interface Raise Event: Connected()                                                    '
Public Sub Connected(ByVal Index As Long)
  Dim oMod As clsModule
  On Error Resume Next
  For Each oMod In mCol
    oMod.Connected Index
  Next oMod
End Sub

