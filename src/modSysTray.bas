Attribute VB_Name = "modSysTray"
Option Explicit

Private Const NOTIFYICONDATA_V1_SIZE As Long = 88
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504

Private Const TRAY_ID = &H125

Dim TrayClasses As Long
Dim TrayClass() As PowerDLL2.SysTray

Public Sub AddTrayClass(NewClass As PowerDLL2.SysTray)
  If (TrayClasses Mod 10) = 0 Then
    ReDim Preserve TrayClass(1 To (TrayClasses + 10)) As PowerDLL2.SysTray
  End If
  TrayClasses = (TrayClasses + 1)
  Set TrayClass(TrayClasses) = NewClass
  TrayClass(TrayClasses).TrayID = TrayClasses
End Sub

Public Sub RemoveTrayClass(RemoveClass As PowerDLL2.SysTray)
  Dim i As Long
  Dim j As Long
  j = GetTrayClassIndexFromID(RemoveClass.TrayID)
  If Not j = 0 Then
    For i = j To (TrayClasses - 1)
      Set TrayClass(i) = TrayClass(i + 1)
    Next
    Set TrayClass(TrayClasses) = Nothing
    TrayClasses = (TrayClasses - 1)
    If (TrayClasses Mod 10) = 0 Then
      If TrayClasses = 0 Then
        Erase TrayClass
      Else
        ReDim Preserve TrayClass(1 To TrayClasses) As PowerDLL2.SysTray
      End If
    End If
  End If
End Sub

Public Function GetSysTraySize() As Long
  Select Case True
    Case IsShellVersion(6)
      GetSysTraySize = NOTIFYICONDATA_V3_SIZE
    Case IsShellVersion(5)
      GetSysTraySize = NOTIFYICONDATA_V2_SIZE
    Case Else
      GetSysTraySize = NOTIFYICONDATA_V1_SIZE
  End Select
End Function

Private Function GetTrayClassIndexFromID(TrayID As Long) As Long
  Dim i As Long
  For i = 1 To TrayClasses
    If TrayClass(i).TrayID = TrayID Then
      GetTrayClassIndexFromID = i
      Exit For
    End If
  Next
End Function

Private Function GetTrayClassIndexFromhWnd(hWndParent As Long) As Long
  Dim i As Long
  For i = 1 To TrayClasses
    If TrayClass(i).hWndParent = hWndParent Then
      GetTrayClassIndexFromhWnd = i
      Exit For
    End If
  Next
End Function

Public Function TrayProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim i As Long
  i = GetTrayClassIndexFromhWnd(hWnd)
  If Not i = 0 Then
    Select Case uMsg
      Case WM_MYHOOK
        With TrayClass(i)
          Select Case lParam
            Case WM_MOUSEMOVE
              .DoThisEvent eMouseMove
            Case WM_LBUTTONDOWN
              .DoThisEvent eLeftMouseDown
            Case WM_LBUTTONUP
              .DoThisEvent eLeftMouseUp
            Case WM_LBUTTONDBLCLK
              .DoThisEvent eLeftMouseDblClick
            Case WM_MBUTTONDOWN
              .DoThisEvent eMiddleMouseDown
            Case WM_MBUTTONUP
              .DoThisEvent eMiddleMouseUp
            Case WM_MBUTTONDBLCLK
              .DoThisEvent eMiddleMouseDblClick
            Case WM_RBUTTONDOWN
              .DoThisEvent eRightMouseDown
            Case WM_RBUTTONUP
              .DoThisEvent eRightMouseUp
            Case WM_RBUTTONDBLCLK
              .DoThisEvent eRightMouseDblClick
            Case WM_MOUSEHOVER
              .DoThisEvent eMouseHover
            Case WM_MOUSELEAVE
              .DoThisEvent eMouseLeave
            Case NIN_BALLOONUSERCLICK
              .DoThisEvent eBalloonTipClicked
            Case NIN_BALLOONTIMEOUT
              .DoThisEvent eBalloonTipTimedOut
          End Select
        End With
      Case Else
        TrayProc = CallWindowProc(TrayClass(i).OldWindowProc, hWnd, uMsg, wParam, lParam)
    End Select
  End If
End Function
