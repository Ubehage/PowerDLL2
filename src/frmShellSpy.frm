VERSION 5.00
Object = "{AD9E813C-FCE6-11D3-8ED3-00E07D815373}#1.0#0"; "MBShSpy.ocx"
Begin VB.Form frmShellSpy 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MBShellSpy.ShellSpy ShellSpy1 
      Left            =   255
      Top             =   330
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "frmShellSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub EnableShellSpy()
  With ShellSpy1
    .DriveEvents = True
    .FolderEvents = True
    .FolderToWatch = mbDesktop
    .ItemEvents = True
    .MediaEvents = True
    .NetworkEvents = True
    .WatchSubFolders = True
    .Enabled = True
  End With
End Sub

Public Sub DisableShellSpy()
  ShellSpy1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  DisableShellSpy
End Sub

Private Sub ShellSpy1_ShellUpdated(ByVal EventID As MBShellSpy.mbShellEventConstants, ByVal DisplayName As String, ByVal Path As String, ByVal PIDL As Long, ByVal DisplayName2 As String, ByVal Path2 As String, ByVal PIDL2 As Long)
  modShellSpy.ShellSpyEventCaller EventID, DisplayName, Path, PIDL, DisplayName2, Path2, PIDL2
End Sub
