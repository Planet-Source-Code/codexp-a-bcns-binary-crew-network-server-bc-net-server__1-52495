VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------'
' clsTokens Class Copyright (C)2003 by CodeXP                     CodeXP@Lycos.de '
'---------------------------------------------------------------------------------'
' YOU CAN USE, YOU CAN CHANGE IT, BUT GIVE IT OTHERS AS IS (WITHOUT CHANGES)      '
'---------------------------------------------------------------------------------'
Option Explicit

Private m_Data        As String
Private m_Limit       As Long
Private m_Empty       As Boolean
Private m_Array()     As String
Private m_Delimiter   As String
Private m_Collection  As Collection
Private b_Changed     As Boolean


Public Property Get Data() As String
Attribute Data.VB_UserMemId = 0
  If b_Changed Then Stringize
  Data = m_Data
End Property


Public Property Let Data(m_New As String)
  m_Data = m_New
  Tokenize
End Property


Public Property Get TokensArray()
  If b_Changed Then
    Stringize
    m_Array = Split(m_Data, m_Delimiter, m_Limit)
  End If
  TokensArray = m_Array
End Property


Public Property Get Count() As Long
  Count = m_Collection.Count
End Property


Public Property Get Token(ByVal Index As Long) As String
  Token = m_Collection(Index)
End Property


Public Property Let Token(ByVal Index As Long, ByVal Value As String)
  Dim i As Long
  If m_Collection(Index) <> Value Then
    b_Changed = True
    If m_Empty Or (Len(Value) > 0) Then
      m_Collection.Add Value, , Index
      i = 1
    End If
    m_Collection.Remove Index + i
  End If
End Property


Public Property Get Delimiter() As String
  Delimiter = m_Delimiter
End Property


Public Function Init(ByVal Data As String, Optional ByVal Delimiter As String = " ", Optional ByVal MaxTokens As Long = -1, Optional ByVal EmptyTokens As Boolean)
  m_Data = Data
  m_Limit = IIf(MaxTokens > 1, MaxTokens, -1)
  m_Empty = EmptyTokens
  m_Delimiter = IIf(Len(Delimiter), Delimiter, " ")
  Tokenize
End Function


Public Sub Remove(ByVal Index As Long)
  b_Changed = True
  m_Collection.Remove Index
End Sub


Public Sub AddToken(ByVal NewToken As String, Optional ByVal Before As Long, Optional ByVal After As Long)
  If m_Empty Or (Len(NewToken) > 0) Then
    b_Changed = True
    If Before > 0 Then
      m_Collection.Add NewToken, , Before
    ElseIf After > 0 Then
      m_Collection.Add NewToken, , , After
    Else
      m_Collection.Add NewToken
    End If
  End If
End Sub


Private Sub Tokenize()
  Dim i As Long
  Set m_Collection = New Collection
  m_Array = Split(m_Data, m_Delimiter, m_Limit)
  For i = 0 To UBound(m_Array)
    If m_Empty Or (Len(m_Array(i)) > 0) Then
      m_Collection.Add m_Array(i)
    End If
  Next i
  b_Changed = False
End Sub


Private Sub Stringize()
  Dim i As Long
  m_Data = ""
  For i = 1 To m_Collection.Count
    m_Data = m_Data & m_Collection(i)
    If i < m_Collection.Count Then
      m_Data = m_Data & m_Delimiter
    End If
  Next i
  b_Changed = False
End Sub


Private Sub Class_Initialize()
  Set m_Collection = New Collection
  m_Delimiter = " " ' set standard Delimiter (white CHR[32])  '
  b_Changed = True  ' for first initialisation  '
End Sub


Private Sub Class_Terminate()
  Set m_Collection = Nothing
End Sub
