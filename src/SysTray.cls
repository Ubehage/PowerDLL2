VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Icon As IPictureDisp
Dim m_Tip As String
Dim m_hWndParent As Long
Dim m_InTray As Boolean

Dim m_TrayID As Long
Dim m_OldWindowProc As Long

Dim niData As NOTIFYICONDATA

Public Event MouseMove()
Public Event LeftMouseDown()
Public Event LeftMouseUp()
Public Event LeftMouseDblClick()
Public Event MiddleMouseDown()
Public Event MiddleMouseUp()
Public Event MiddleMouseDblClick()
Public Event RightMouseDown()
Public Event RightMouseUp()
Public Event RightMouseDblClick()
Public Event MouseHover()
Public Event MouseLeave()
Public Event BalloonTipClick()
Public Event BalloonTipTimeout()

Public Property Get Icon() As IPictureDisp
  Set Icon = m_Icon
End Property
Public Property Get Tip() As String
  Tip = m_Tip
End Property
Public Property Get hWndParent() As Long
  hWndParent = m_hWndParent
End Property
Public Property Get InTray() As Boolean
  InTray = m_InTray
End Property
Friend Property Get TrayID() As Long
  TrayID = m_TrayID
End Property
Friend Property Get OldWindowProc() As Long
  OldWindowProc = m_OldWindowProc
End Property
Public Property Let Icon(New_Icon As IPictureDisp)
  If Not m_Icon Is New_Icon Then
    Set m_Icon = New_Icon
    If m_InTray Then
      If m_Icon Is Nothing Then
        InTray = False
      Else
        ChangeTrayIcon
      End If
    End If
  End If
End Property
Public Property Let Tip(New_Tip As String)
  If Not m_Tip = New_Tip Then
    m_Tip = New_Tip
    If m_InTray Then
      ChangeTrayTip
    End If
  End If
End Property
Public Property Let hWndParent(New_hWndParent As Long)
  If Not m_hWndParent = New_hWndParent Then
    If m_InTray Then
      If RemoveFromTray Then
        m_hWndParent = New_hWndParent
        If Not m_hWndParent = 0 Then
          If Not PutInTray Then
            m_InTray = False
          End If
        Else
          m_InTray = False
        End If
      End If
    Else
      m_hWndParent = New_hWndParent
    End If
  End If
End Property
Public Property Let InTray(New_InTray As Boolean)
  If Not m_InTray = New_InTray Then
    If m_InTray Then
      If RemoveFromTray Then
        m_InTray = False
      End If
    Else
      If PutInTray Then
        m_InTray = True
      End If
    End If
  End If
End Property
Friend Property Let TrayID(New_TrayID As Long)
  m_TrayID = New_TrayID
End Property
Friend Property Let OldWindowProc(New_OldWindowProc As Long)
  m_OldWindowProc = New_OldWindowProc
End Property

Public Sub ShowBalloonTip(Title As String, Text As String)
  If m_InTray Then
    With niData
      .uFlags = NIF_INFO
      .uTimeoutAndVersion = 30000
      .dwInfoFlags = NIIF_INFO
      .szInfoTitle = Title
      .szInfo = Text
    End With
    Shell_NotifyIcon NIM_MODIFY, niData
  End If
End Sub

Friend Sub DoThisEvent(EventToDo As EVENT_Constants)
  Select Case EventToDo
    Case EVENT_Constants.eMouseMove
      RaiseEvent MouseMove
    Case EVENT_Constants.eLeftMouseDown
      RaiseEvent LeftMouseDown
    Case EVENT_Constants.eLeftMouseUp
      RaiseEvent LeftMouseUp
    Case EVENT_Constants.eLeftMouseDblClick
      RaiseEvent LeftMouseDblClick
    Case EVENT_Constants.eMiddleMouseDown
      RaiseEvent MiddleMouseDown
    Case EVENT_Constants.eMiddleMouseUp
      RaiseEvent MiddleMouseUp
    Case EVENT_Constants.eMiddleMouseDblClick
      RaiseEvent MiddleMouseDblClick
    Case EVENT_Constants.eRightMouseDown
      RaiseEvent RightMouseDown
    Case EVENT_Constants.eRightMouseUp
      RaiseEvent RightMouseUp
    Case EVENT_Constants.eRightMouseDblClick
      RaiseEvent RightMouseDblClick
    Case EVENT_Constants.eMouseHover
      RaiseEvent MouseHover
    Case EVENT_Constants.eMouseLeave
      RaiseEvent MouseLeave
    Case EVENT_Constants.eBalloonTipClicked
      RaiseEvent BalloonTipClick
    Case EVENT_Constants.eBalloonTipTimedOut
      RaiseEvent BalloonTipTimeout
  End Select
End Sub

Private Function RegisterWindowProc() As Boolean
  If Not m_hWndParent = 0 Then
    m_OldWindowProc = SetWindowLong(m_hWndParent, GWL_WNDPROC, AddressOf TrayProc)
    RegisterWindowProc = True
  End If
End Function

Private Function UnRegisterWindowProc() As Boolean
  If Not m_hWndParent = 0 Then
    SetWindowLong m_hWndParent, GWL_WNDPROC, m_OldWindowProc
    m_OldWindowProc = 0
    UnRegisterWindowProc = True
  End If
End Function

Private Function PutInTray() As Boolean
  If Not m_InTray Then
    If Not m_Icon Is Nothing Then
      If RegisterWindowProc Then
        With niData
          .cbSize = GetSysTraySize
          .hWnd = m_hWndParent
          .uID = m_TrayID
          .uFlags = NIF_ICON Or NIF_MESSAGE
          If Not m_Tip = "" Then
            .uFlags = .uFlags Or NIF_TIP
            .szTip = (m_Tip + vbNullChar)
          End If
          .uCallbackMessage = WM_MYHOOK
          .hIcon = m_Icon
        End With
        Shell_NotifyIcon NIM_ADD, niData
        PutInTray = True
      End If
    End If
  End If
End Function

Private Function RemoveFromTray() As Boolean
  If m_InTray Then
    If Not m_hWndParent = 0 Then
      If UnRegisterWindowProc Then
        Shell_NotifyIcon NIM_DELETE, niData
        RemoveFromTray = True
      End If
    End If
  End If
End Function

Private Sub ChangeTrayIcon()
  If m_InTray Then
    With niData
      .uFlags = NIF_ICON
      .hIcon = m_Icon
    End With
    Shell_NotifyIcon NIM_MODIFY, niData
  End If
End Sub

Private Sub ChangeTrayTip()
  If m_InTray Then
    With niData
      .uFlags = NIF_TIP
      .szTip = (m_Tip + vbNullChar)
    End With
    Shell_NotifyIcon NIM_MODIFY, niData
  End If
End Sub

Private Sub Class_Initialize()
  AddTrayClass Me
End Sub

Private Sub Class_Terminate()
  If m_InTray Then
    RemoveFromTray
  End If
  RemoveTrayClass Me
End Sub
