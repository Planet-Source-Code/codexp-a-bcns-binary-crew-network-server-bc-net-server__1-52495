Attribute VB_Name = "modTime"
'---------------------------------------------------------------------------------'
' modTime (Time Functions) Copyright (C)2003 by CodeXP            CodeXP@Lycos.de '
'---------------------------------------------------------------------------------'
' YOU CAN USE, YOU CAN CHANGE IT, BUT GIVE IT OTHERS AS IS (WITHOUT CHANGES)      '
'---------------------------------------------------------------------------------'
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Function TickToDays(ByVal Tick As Long) As Long
  Tick = Int(Tick / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 24)
  TickToDays = Tick
End Function


Public Function TickToHours(ByVal Tick As Long) As Long
  Tick = Int(Tick / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  TickToHours = Tick
End Function


Public Function TickToMinutes(ByVal Tick As Long) As Long
  Tick = Int(Tick / 1000)
  Tick = Int(Tick / 60)
  TickToMinutes = Tick
End Function


Public Function TickToSeconds(ByVal Tick As Long) As Long
  TickToSeconds = Int(Tick / 1000)
End Function


Public Function TickDays() As Long
  Dim Tick As Long
  Tick = Int(GetTickCount / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 24)
  TickDays = Tick
End Function


Public Function TickHours() As Long
  Dim Tick As Long
  Tick = Int(GetTickCount / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  TickHours = Tick
End Function


Public Function TickMinutes() As Long
  Dim Tick As Long
  Tick = Int(GetTickCount / 1000)
  Tick = Int(Tick / 60)
  TickMinutes = Tick
End Function


Public Function TickSeconds() As Long
  TickSeconds = CLng(GetTickCount / 1000)
End Function


Public Function TickMilli() As Long
  TickMilli = CLng(GetTickCount / 100)
End Function


Public Function Tick() As Long
  Tick = GetTickCount
End Function


Public Function TimeLong() As Double
  TimeLong = TimeToLong(Now)
End Function


Public Function TimeToLong(TimeStamp As String) As Double
  TimeToLong = Int(CDbl(CDate(TimeStamp)) * 100000)
End Function


Public Function LongToTime(LongStamp As Double) As Date
  LongToTime = CDate(LongStamp / 100000)
End Function


Public Function TimeStamp() As String
  TimeStamp = TimeToStamp(TimeToLong(Now))
End Function


Public Function TimeToStamp(LongStamp As Double, Optional Fmt As String) As String
  Dim lDays As Long
  Dim lHours As Long
  Select Case UCase(Trim(Fmt))
    Case "TIME"
      lDays = Val(Format(LongToTime(LongStamp), "d"))
      lHours = Val(Format(LongToTime(LongStamp), "h")) + lDays * 24
      TimeToStamp = lHours & Format(LongToTime(LongStamp), ":mm:ss")
    Case Else
      TimeToStamp = Format(LongToTime(LongStamp), "\[dd.mm.yy hh:mm:ss\] ")
  End Select
End Function


Public Function TimeLeftAfter(LongStamp As Double) As String
  Dim lH As Long
  Dim lM As Long
  Dim lS As Long
  lS = SecondsLeftAfter(LongStamp) Mod 60
  lM = MinutesLeftAfter(LongStamp) Mod 60
  lH = HoursLeftAfter(LongStamp)
  TimeLeftAfter = lH & ":" & IIf(lM < 10, "0", "") & lM & ":" & IIf(lS < 10, "0", "") & lS
End Function


Public Function SecondsLeftAfter(LongStamp As Double) As Long
  On Error Resume Next
  SecondsLeftAfter = DateDiff("s", LongToTime(LongStamp), Now)
End Function


Public Function MinutesLeftAfter(LongStamp As Double) As Long
  On Error Resume Next
  MinutesLeftAfter = DateDiff("s", LongToTime(LongStamp), Now) / 60
End Function


Public Function HoursLeftAfter(LongStamp As Double) As Long
  On Error Resume Next
  HoursLeftAfter = DateDiff("s", LongToTime(LongStamp), Now) / 60 / 60
End Function


Public Function YearsLeftAfter(LongStamp As Double) As Long
  On Error Resume Next
  YearsLeftAfter = Val(DateDiff("yyyy", LongToTime(LongStamp), Now))
End Function


