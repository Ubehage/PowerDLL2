Attribute VB_Name = "modShellChange3"
Option Explicit

Public Function ShellChangeWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim shClass As clsShellChange
  Set shClass = GetShellChangeClassFromCallbackMessage(uMsg)
  If Not shClass Is Nothing Then
    shClass.ShellChangeNotificationEvent wParam, lParam
  End If
  ShellChangeWindowProc = CallWindowProc(GetProp(hWnd, OLDWNDPROCNAME), hWnd, uMsg, wParam, lParam)
End Function

Public Function SubClassShellChangeClass(shClass As clsShellChange) As Boolean
  If Not shClass.OldWindowProc = 0 Then
    If Not UnsubClassShellChangeClass(shClass) Then
      Exit Function
    End If
  End If
  shClass.OldWindowProc = SetWindowLong(shClass.hWndParent, GWL_WNDPROC, AddressOf ShellChangeWindowProc)
  If shClass.OldWindowProc Then
    SubClassShellChangeClass = SetProp(shClass.hWndParent, OLDWNDPROCNAME, shClass.OldWindowProc)
  End If
End Function

Public Function UnsubClassShellChangeClass(shClass As clsShellChange) As Boolean
  If shClass.OldWindowProc Then
    UnsubClassShellChangeClass = SetWindowLong(shClass.hWndParent, GWL_WNDPROC, shClass.OldWindowProc)
    If UnsubClassShellChangeClass Then
      shClass.OldWindowProc = 0
      RemoveProp shClass.hWndParent, OLDWNDPROCNAME
    End If
  End If
End Function

Public Function SHNotify_Register(shClass As clsShellChange) As Boolean
  Dim ps As PIDLSTRUCT
  If shClass.NotifyID = 0 Then
    shClass.PIDLDesktop = GetPIDLFromFolderID(0, CSIDL_DESKTOP)
    If shClass.PIDLDesktop Then
      ps.pidl = shClass.PIDLDesktop
      ps.bWatchSubFolders = True
      shClass.NotifyID = SHChangeNotifyRegister(shClass.hWndParent, nfType Or nfIDList, scAllEvents Or scInterrupt, shClass.CallbackMessage, 1, ps)
      SHNotify_Register = CBool(shClass.NotifyID)
    Else
      CoTaskMemFree shClass.PIDLDesktop
    End If
  End If
End Function

Public Function SHNotify_Unregister(shClass As clsShellChange) As Boolean
  If shClass.NotifyID Then
    If SHChangeNotifyDeregister(shClass.NotifyID) Then
      shClass.NotifyID = 0
      CoTaskMemFree shClass.PIDLDesktop
      shClass.PIDLDesktop = 0
      SHNotify_Unregister = True
    End If
  End If
End Function
