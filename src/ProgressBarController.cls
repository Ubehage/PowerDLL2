VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgressBarController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_ProgressBar As ComctlLib.ProgressBar
Dim m_Min As Double
Dim m_Max As Double
Dim m_Value As Double
Dim m_FloodPercent As Single

Public Property Get ProgressBar() As Object
  Set ProgressBar = m_ProgressBar
End Property
Public Property Get Min() As Double
  If Not m_ProgressBar Is Nothing Then
    m_Min = m_ProgressBar.Min
  End If
  Min = m_Min
End Property
Public Property Get Max() As Double
  If Not m_ProgressBar Is Nothing Then
    m_Max = m_ProgressBar.Max
  End If
  Max = m_Max
End Property
Public Property Get Value() As Double
  If Not m_ProgressBar Is Nothing Then
    m_Value = m_ProgressBar.Value
  End If
  Value = m_Value
End Property
Public Property Get FloodPercent() As Single
  FloodPercent = CalculateFloodPercent
End Property
Public Property Let ProgressBar(New_ProgressBar As Object)
  If TypeName(New_ProgressBar) = "ProgressBar" Then
    If Not m_ProgressBar Is New_ProgressBar Then
      Set m_ProgressBar = New_ProgressBar
    End If
  End If
End Property
Public Property Let Min(New_Min As Double)
  If Not m_Min = New_Min Then
    If New_Min > m_Value Then
      If New_Min > m_Max Then
        m_Min = m_Value
      Else
        m_Value = New_Min
        m_Min = New_Min
      End If
    Else
      m_Min = New_Min
    End If
    Refresh
  End If
End Property
Public Property Let Max(New_Max As Double)
  If Not m_Max = New_Max Then
    If New_Max < m_Value Then
      If New_Max < m_Min Then
        m_Max = m_Value
      Else
        m_Value = New_Max
        m_Max = New_Max
      End If
    Else
      m_Max = New_Max
    End If
    Refresh
  End If
End Property
Public Property Let Value(New_Value As Double)
  If Not m_Value = New_Value Then
    If New_Value < m_Min Then
      m_Value = m_Min
    ElseIf New_Value > m_Max Then
      m_Value = m_Max
    Else
      m_Value = New_Value
    End If
    Refresh
  End If
End Property
Public Property Let FloodPercent(New_FloodPercent As Single)
  If Not m_FloodPercent = New_FloodPercent Then
    If (New_FloodPercent >= 0 And New_FloodPercent <= 100) Then
      m_FloodPercent = New_FloodPercent
      m_Min = Min
      m_Max = Max
      m_Value = (m_Min + (((m_Max - m_Min) / 100) * m_FloodPercent))
    End If
  End If
  Refresh
End Property

Public Sub Refresh()
  If Not m_ProgressBar Is Nothing Then
    With m_ProgressBar
      .Min = m_Min
      .Max = m_Max
      .Value = m_Value
    End With
  End If
End Sub

Private Function CalculateFloodPercent() As Single
  m_Min = Min
  m_Max = Max
  m_Value = Value
  CalculateFloodPercent = ((100 / (m_Max - m_Min)) * (m_Value - m_Min))
End Function

Private Sub Class_Terminate()
  Set m_ProgressBar = Nothing
  m_Min = 0
  m_Max = 0
  m_Value = 0
  m_FloodPercent = 0
End Sub
