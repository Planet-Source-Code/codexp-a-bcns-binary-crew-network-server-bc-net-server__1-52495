VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsUser"
Attribute VB_Ext_KEY = "Member0" ,"clsUser"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mCol  As Collection
Private SAI   As clsNetAppInterface


Public Sub Init(SAIRef As clsNetAppInterface)
  Set SAI = SAIRef
End Sub


Public Function LoginUser(ByVal Index As Long, ByVal sUser As String, Optional ByVal sPass As String) As Boolean
  Dim oNewUser As clsUser
  
  On Error Resume Next
  
  sPass = Trim(sPass)
  sUser = Trim(sUser)
  
  If UCase(sUser) = "SERVER" Then Exit Function
  Set oNewUser = New clsUser
  Call oNewUser.Init(SAI, Index, sUser)
  mCol.Add oNewUser, oNewUser.UserID
  
  LoginUser = True
LoginUser_Error:
  Set oNewUser = Nothing
End Function


Public Property Get User(IndexKey) As clsUser
Attribute User.VB_UserMemId = 0
  Set User = mCol(IndexKey)
End Property


Public Property Get UsersCount() As Long
  UsersCount = mCol.Count
End Property


Public Sub Remove(IndexKey)
  mCol.Remove IndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub


Public Sub SendToAll(ByVal CmdLine As String, Optional ByVal ExceptIndex As Long = -1, Optional ByVal ExceptUserID As String)
  Dim oUsr As clsUser
  For Each oUsr In mCol
    If (oUsr.Index <> ExceptIndex) And (oUsr.UserID <> ExceptUserID) Then
      SAI.SendLine oUsr.Index, CmdLine
    End If
  Next oUsr
End Sub


Public Sub SendUserList(ByVal Index As Long, Optional ByVal ExceptIndex As Long = -1)
  Dim oUsr  As clsUser
  For Each oUsr In mCol
    If oUsr.Index <> ExceptIndex Then
      SAI.SendLine Index, ":" & oUsr.UserID & "@" & SAI.HostName & " Comm.Login"
    End If
  Next oUsr
End Sub


Public Function UserExists(IndexKey) As Boolean
  If VarType(IndexKey) = vbString Then
    If UCase(IndexKey) = "ADMIN" Then
      UserExists = True
      Exit Function
    End If
  End If
  
  Dim oUsr As clsUser
  On Error Resume Next
  Set oUsr = mCol(IndexKey)
  UserExists = (Err = 0)
End Function


Public Function GetUserID(ByVal Index As Long) As String
  Dim oUsr As clsUser
  For Each oUsr In mCol
    If oUsr.Index = Index Then
      GetUserID = oUsr.UserID
      Exit For
    End If
  Next oUsr
End Function

