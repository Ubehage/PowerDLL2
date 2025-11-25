Attribute VB_Name = "modTimer"
Option Explicit

Private Const WM_TIMER As Long = &H113

Dim TimerClasses As Long
Dim TimerClass() As PowerDLL2.Timer

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal UElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Sub AddTimerClass(NewClass As PowerDLL2.Timer)
  If (TimerClasses Mod 10) = 0 Then
    ReDim Preserve TimerClass(1 To (TimerClasses + 10)) As PowerDLL2.Timer
  End If
  TimerClasses = (TimerClasses + 1)
  Set TimerClass(TimerClasses) = NewClass
  TimerClass(TimerClasses).TimerID = TimerClasses
End Sub

Public Sub RemoveTimerClass(RemClass As PowerDLL2.Timer)
  Dim i As Long
  Dim j As Long
  j = GetTimerIndexFromClass(RemClass)
  For i = j To (TimerClasses - 1)
    Set TimerClass(i) = TimerClass((i + 1))
  Next
  Set TimerClass(TimerClasses) = Nothing
  TimerClasses = (TimerClasses - 1)
  If (TimerClasses Mod 10) = 0 Then
    If TimerClasses = 0 Then
      Erase TimerClass
    Else
      ReDim Preserve TimerClass(1 To TimerClasses) As PowerDLL2.Timer
    End If
  End If
End Sub

Public Function StartTimer(StartClass As PowerDLL2.Timer) As Boolean
  If Not StartClass.Interval = 0 Then
    If Not StartClass.hWndParent = 0 Then
      If Not StartClass.Enabled Then
        SetTimer StartClass.hWndParent, StartClass.TimerID, StartClass.Interval, AddressOf TimerProc
        StartTimer = True
      End If
    End If
  End If
End Function

Public Sub StopTimer(StopClass As PowerDLL2.Timer)
  If StopClass.Enabled Then
    KillTimer StopClass.hWndParent, StopClass.TimerID
  End If
End Sub

Private Function GetTimerIndexFromClass(tClass As PowerDLL2.Timer) As Long
  GetTimerIndexFromClass = GetTimerIndexFromID(tClass.TimerID)
End Function

Private Function GetTimerIndexFromID(TimerID As Long) As Long
  Dim i As Long
  For i = 1 To TimerClasses
    If TimerClass(i).TimerID = TimerID Then
      GetTimerIndexFromID = i
      Exit For
    End If
  Next
End Function

Public Function TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
  If uMsg = WM_TIMER Then
    TimerClass(GetTimerIndexFromID(idEvent)).DoTimer
  End If
End Function
