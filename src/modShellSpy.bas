Attribute VB_Name = "modShellSpy"
Option Explicit

Global ShellSpyFormLoaded As Boolean

Dim ReggedShellSpys As Long
Dim ReggedShellSpy() As PowerDLL2.ShellSpy

Public Function RegisterShellSpyClass(ShellSpyClass As PowerDLL2.ShellSpy) As Boolean
  If Not IsShellSpyClassRegistered(ShellSpyClass) Then
    If (ReggedShellSpys Mod 10) = 0 Then
      ReDim Preserve ReggedShellSpy(1 To (ReggedShellSpys + 10)) As PowerDLL2.ShellSpy
    End If
    ReggedShellSpys = (ReggedShellSpys + 1)
    Set ReggedShellSpy(ReggedShellSpys) = ShellSpyClass
    LoadShellSpyForm
    RegisterShellSpyClass = True
  End If
End Function

Public Function UnRegisterShellSpyClass(ShellSpyClass As PowerDLL2.ShellSpy) As Boolean
  Dim i As Long
  Dim sIndex As Long
  sIndex = GetShellSpyClassIndex(ShellSpyClass)
  If Not sIndex = 0 Then
    For i = sIndex To (ReggedShellSpys - 1)
      Set ReggedShellSpy(i) = ReggedShellSpy((i + 1))
    Next
    Set ReggedShellSpy(ReggedShellSpys) = Nothing
    ReggedShellSpys = (ReggedShellSpys - 1)
    If (ReggedShellSpys Mod 10) = 0 Then
      If ReggedShellSpys = 0 Then
        Erase ReggedShellSpy
      Else
        ReDim Preserve ReggedShellSpy(1 To ReggedShellSpys) As PowerDLL2.ShellSpy
      End If
    End If
    UnloadShellSpyForm
    UnRegisterShellSpyClass = True
  End If
End Function

Public Sub ShellSpyEventCaller(EventID As mbShellEventConstants, DisplayName As String, Path As String, PIDL As Long, DisplayName2 As String, Path2 As String, PIDL2 As Long)
  Dim i As Long
  For i = 1 To ReggedShellSpys
    If ReggedShellSpy(i).Enabled Then
      ReggedShellSpy(i).EventCaller EventID, DisplayName, Path, PIDL, DisplayName2, Path2, PIDL2
    End If
  Next
End Sub

Private Function IsShellSpyClassRegistered(ShellSpyClass As PowerDLL2.ShellSpy) As Boolean
  IsShellSpyClassRegistered = Not (GetShellSpyClassIndex(ShellSpyClass) = 0)
End Function

Private Function GetShellSpyClassIndex(ShellSpyClass As PowerDLL2.ShellSpy) As Long
  Dim i As Long
  For i = 1 To ReggedShellSpys
    If ShellSpyClass Is ReggedShellSpy(i) Then
      GetShellSpyClassIndex = i
      Exit For
    End If
  Next
End Function

Private Sub LoadShellSpyForm()
  If Not ShellSpyFormLoaded Then
    If ReggedShellSpys > 0 Then
      Load frmShellSpy
      frmShellSpy.EnableShellSpy
      ShellSpyFormLoaded = True
    End If
  End If
End Sub

Private Sub UnloadShellSpyForm()
  If ShellSpyFormLoaded Then
    If ReggedShellSpys = 0 Then
      frmShellSpy.DisableShellSpy
      Unload frmShellSpy
      Set frmShellSpy = Nothing
      ShellSpyFormLoaded = False
    End If
  End If
End Sub
