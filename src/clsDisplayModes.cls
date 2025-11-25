VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisplayModes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const BITSPIXEL = 12
Private Const HORZRES = 8
Private Const VERTRES = 10

Dim m_Count As Long
Dim m_DisplayMode() As PowerDLL2.clsDisplayMode

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Property Get Count() As Long
  Count = m_Count
End Property
Public Property Get DisplayMode(Index As Long) As PowerDLL2.clsDisplayMode
  Set DisplayMode = m_DisplayMode(Index)
End Property

Public Function GetDisplayMode(Width As Long, Height As Long, Bits As Long) As PowerDLL2.clsDisplayMode
  Dim i As Long
  i = GetDisplayModeIndex(Width, Height, Bits)
  If Not i = 0 Then
    Set GetDisplayMode = m_DisplayMode(i)
  End If
End Function

Public Function GetCurrentDisplayMode() As PowerDLL2.clsDisplayMode
  Dim hWndDesk As Long
  Dim hDcDesk As Long
  Dim sWidth As Long
  Dim sHeight As Long
  Dim sBits As Long
  hWndDesk = GetDesktopWindow
  hDcDesk = GetWindowDC(hWndDesk)
  sWidth = GetDeviceCaps(hDcDesk, HORZRES)
  sHeight = GetDeviceCaps(hDcDesk, VERTRES)
  sBits = GetDeviceCaps(hDcDesk, BITSPIXEL)
  ReleaseDC hWndDesk, hDcDesk
  Set GetCurrentDisplayMode = GetDisplayMode(sWidth, sHeight, sBits)
End Function

Private Sub Refresh()
  Clear
  GetDisplayModes
End Sub

Private Sub Clear()
  m_Count = 0
  Erase m_DisplayMode
End Sub

Private Function GetDisplayModeIndex(Width As Long, Height As Long, Bits As Long) As Long
  Dim i As Long
  For i = 1 To m_Count
    If ((m_DisplayMode(i).Width = Width And m_DisplayMode(i).Height = Height) And m_DisplayMode(i).Bits = Bits) Then
      GetDisplayModeIndex = i
      Exit For
    End If
  Next
End Function

Private Sub AddDisplayMode(Width As Long, Height As Long, Bits As Long)
  If (m_Count Mod 10) = 0 Then
    ReDim Preserve m_DisplayMode(1 To (m_Count + 10)) As PowerDLL2.clsDisplayMode
  End If
  m_Count = (m_Count + 1)
  Set m_DisplayMode(m_Count) = New clsDisplayMode
  With m_DisplayMode(m_Count)
    .ModesParent = Me
    .Width = Width
    .Height = Height
    .Bits = Bits
  End With
End Sub

Private Sub GetDisplayModes()
  Dim dMode As DEVMODE
  Dim dM As Long
  Dim r As Long
  dMode.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
  dMode.dmSize = LenB(dMode)
  dM = 0
  Do While EnumDisplaySettings(0&, dM, dMode) > 0
    If dMode.dmBitsPerPel > 4 Then
      AddNewDisplayMode dMode
    End If
    dM = (dM + 1)
  Loop
End Sub

Private Sub AddNewDisplayMode(NewMode As DEVMODE)
  Dim nWidth As Long
  Dim nHeight As Long
  Dim nBits As Long
  With NewMode
    nWidth = .dmPelsWidth
    nHeight = .dmPelsHeight
    nBits = .dmBitsPerPel
  End With
  If Not IsDisplayModeAdded(nWidth, nHeight, nBits) Then
    AddDisplayMode nWidth, nHeight, nBits
  End If
End Sub

Private Function IsDisplayModeAdded(Width As Long, Height As Long, Bits As Long) As Boolean
  IsDisplayModeAdded = Not (GetDisplayModeIndex(Width, Height, Bits) = 0)
End Function

Private Sub Class_Initialize()
  Refresh
End Sub

Private Sub Class_Terminate()
  Clear
End Sub
