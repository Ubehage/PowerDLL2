Attribute VB_Name = "modShellChange2"
Option Explicit

Private Const WM_SHNOTIFY = &H401

Dim ShellChangeClasses As Long
Dim ShellChangeClass() As clsShellChange

Public Sub RegisterShellChangeClass(shClass As clsShellChange)
  If Not IsShellChangeClassRegistered(shClass) Then
    If (ShellChangeClasses Mod 10) = 0 Then
      ReDim Preserve ShellChangeClass(1 To (ShellChangeClasses + 10)) As clsShellChange
    End If
    ShellChangeClasses = (ShellChangeClasses + 1)
    Set ShellChangeClass(ShellChangeClasses) = shClass
    ShellChangeClass(ShellChangeClasses).Index = ShellChangeClasses
    ShellChangeClass(ShellChangeClasses).CallbackMessage = GetNewCallbackMessage
  End If
End Sub

Public Sub UnregisterShellChangeClass(shClass As clsShellChange)
  Dim i As Long
  Dim j As Long
  j = GetShellChangeClassIndex(shClass)
  If Not j = 0 Then
    For i = j To (ShellChangeClasses - 1)
      Set ShellChangeClass(i) = ShellChangeClass((i + 1))
      ShellChangeClass(i).Index = i
    Next
    ShellChangeClass(ShellChangeClasses).Index = 0
    Set ShellChangeClass(ShellChangeClasses) = Nothing
    ShellChangeClasses = (ShellChangeClasses - 1)
    If (ShellChangeClasses Mod 10) = 0 Then
      If ShellChangeClasses = 0 Then
        Erase ShellChangeClass
      Else
        ReDim Preserve ShellChangeClass(1 To ShellChangeClasses) As clsShellChange
      End If
    End If
  End If
End Sub

Public Function GetShellChangeClassFromCallbackMessage(CallbackMessage As Long) As clsShellChange
  Dim i As Long
  For i = 1 To ShellChangeClasses
    If ShellChangeClass(i).CallbackMessage = CallbackMessage Then
      Set GetShellChangeClassFromCallbackMessage = ShellChangeClass(i)
      Exit For
    End If
  Next
End Function

Private Function IsShellChangeClassRegistered(shClass As clsShellChange) As Boolean
  IsShellChangeClassRegistered = Not (GetShellChangeClassIndex(shClass) = 0)
End Function

Private Function GetShellChangeClassIndex(shClass As clsShellChange) As Long
  GetShellChangeClassIndex = shClass.Index
End Function

Private Function IsCallbackMessageInUse(CallbackMessage As Long) As Boolean
  IsCallbackMessageInUse = Not (GetShellChangeClassFromCallbackMessage(CallbackMessage) Is Nothing)
End Function

Private Function GetNewCallbackMessage() As Long
  Dim nMessage As Long
  nMessage = WM_SHNOTIFY
  While IsCallbackMessageInUse(nMessage)
    nMessage = (nMessage + 1)
  Wend
  GetNewCallbackMessage = nMessage
End Function
