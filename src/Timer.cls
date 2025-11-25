VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Interval As Long
Dim m_Enabled As Boolean
Dim m_hWndParent As Long

Dim m_TimerID As Long

Public Event Timer()

Public Property Get Interval() As Long
  Interval = m_Interval
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Get hWndParent() As Long
  hWndParent = m_hWndParent
End Property
Friend Property Get TimerID() As Long
  TimerID = m_TimerID
End Property
Public Property Let Interval(New_Interval As Long)
  If Not m_Interval = New_Interval Then
    m_Interval = New_Interval
    If m_Enabled Then
      RestartTimer
    End If
  End If
End Property
Public Property Let Enabled(New_Enabled As Boolean)
  If Not m_Enabled = New_Enabled Then
    If New_Enabled Then
      If StartTimer(Me) Then
        m_Enabled = True
      End If
    Else
      StopTimer Me
      m_Enabled = False
    End If
  End If
End Property
Public Property Let hWndParent(New_hWndParent As Long)
  If Not m_hWndParent = New_hWndParent Then
    If m_Enabled Then
      Enabled = False
    End If
    m_hWndParent = New_hWndParent
  End If
End Property
Friend Property Let TimerID(New_TimerID As Long)
  m_TimerID = New_TimerID
End Property

Public Sub RestartTimer()
  If m_Enabled Then
    Enabled = False
  End If
  Enabled = True
End Sub

Friend Sub DoTimer()
  RaiseEvent Timer
End Sub

Private Sub Class_Initialize()
  AddTimerClass Me
End Sub

Private Sub Class_Terminate()
  Enabled = False
  RemoveTimerClass Me
End Sub
