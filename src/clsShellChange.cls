VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShellChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_hWndParent As Long
Dim m_CallbackMessage As Long
Dim m_Index As Long
Dim m_WindowProc As Long
Dim m_OldWindowProc As Long
Dim m_NotifyID As Long
Dim m_PIDLDesktop As Long

Dim IsReceiving As Boolean

Public Event FreeSpaceChanged(Path As String)
Public Event FileRenamed(OldPath As String, NewPath As String)
Public Event FileChanged(Path As String)
Public Event FileAssociationChanged(Path As String)
Public Event FileDeleted(Path As String)
Public Event FileCreated(Path As String)
Public Event FolderRenamed(OldPath As String, NewPath As String)
Public Event FolderChanged(Path As String)
Public Event FolderAssociationChanged(Path As String)
Public Event FolderDeleted(Path As String)
Public Event FolderCreated(Path As String)

Public Property Get hWndParent() As Long
  hWndParent = m_hWndParent
End Property
Friend Property Get CallbackMessage() As Long
  CallbackMessage = m_CallbackMessage
End Property
Friend Property Get Index() As Long
  Index = m_Index
End Property
Friend Property Get WindowProc() As Long
  WindowProc = m_WindowProc
End Property
Friend Property Get OldWindowProc() As Long
  OldWindowProc = m_OldWindowProc
End Property
Friend Property Get NotifyID() As Long
  NotifyID = m_NotifyID
End Property
Friend Property Get PIDLDesktop() As Long
  PIDLDesktop = m_PIDLDesktop
End Property
Public Property Let hWndParent(New_hWndParent As Long)
  If Not m_hWndParent = New_hWndParent Then
    m_hWndParent = New_hWndParent
  End If
End Property
Friend Property Let CallbackMessage(New_CallbackMessage As Long)
  m_CallbackMessage = New_CallbackMessage
End Property
Friend Property Let Index(New_Index As Long)
  m_Index = New_Index
End Property
Friend Property Let WindowProc(New_WindowProc As Long)
  m_WindowProc = New_WindowProc
End Property
Friend Property Let OldWindowProc(New_OldWindowProc As Long)
  m_OldWindowProc = New_OldWindowProc
End Property
Friend Property Let NotifyID(New_NotifyID As Long)
  m_NotifyID = New_NotifyID
End Property
Friend Property Let PIDLDesktop(New_PIDLDesktop As Long)
  m_PIDLDesktop = New_PIDLDesktop
End Property

Public Function StartReceiving() As Boolean
  If Not IsReceiving Then
    If SubClassShellChangeClass(Me) Then
      IsReceiving = SHNotify_Register(Me)
      StartReceiving = IsReceiving
    End If
  End If
End Function

Public Function StopReceiving() As Boolean
  If IsReceiving Then
    If SHNotify_Unregister(Me) Then
      If UnsubClassShellChangeClass(Me) Then
        IsReceiving = False
        StopReceiving = True
      End If
    End If
  End If
End Function

Friend Sub ShellChangeNotificationEvent(Message1 As Long, Message2 As Long)
  Dim sOut As String
  Dim shns As SHNOTIFYSTRUCT
  CopyMemory shns, ByVal Message1, Len(shns)
  Select Case Message2
    Case SHELLCHANGE_EVENT_IDs.scFreeSpaceChanged
      Dim dwDriveBits As Long
      Dim wHighBit As Integer
      Dim wBit As Integer
      CopyMemory dwDriveBits, ByVal shns.dwItem1 + 2, 4
      wHighBit = Int((Log(dwDriveBits) / Log(2)))
      For wBit = 0 To wHighBit
        If ((2 ^ wBit) And dwDriveBits) Then
          RaiseEvent FreeSpaceChanged((Chr((vbKeyA + wBit)) + ":"))
        End If
      Next
    Case SHELLCHANGE_EVENT_IDs.scIconUpdated
      'Something...
    Case Else
      Dim Item1 As String
      Dim Item2 As String
      Dim scEvent As SHELLCHANGE_EVENT_IDs
      If shns.dwItem1 Then
        SHGetPathFromIDList shns.dwItem1, Item1
      End If
      If shns.dwItem2 Then
        SHGetPathFromIDList shns.dwItem2, Item2
      End If
      scEvent = GetShellChangeEventFromLong(Message2)
      If scEvent = scAssociationChanged Then
        RaiseEvent FileAssociationChanged(Item1)
      ElseIf scEvent = scFileContentChanged Then
        RaiseEvent FileChanged(Item1)
      ElseIf scEvent = scFileCreated Then
        RaiseEvent FileCreated(Item1)
      ElseIf scEvent = scFileDeleted Then
        RaiseEvent FileDeleted(Item1)
      ElseIf scEvent = scFileRenamed Then
        RaiseEvent FileRenamed(Item1, Item2)
      ElseIf scEvent = scFolderContentChanged Then
        RaiseEvent FolderChanged(Item1)
      ElseIf scEvent = scFolderCreated Then
        RaiseEvent FolderCreated(Item1)
      ElseIf scEvent = scFolderDeleted Then
        RaiseEvent FolderDeleted(Item1)
      ElseIf scEvent = scFolderRenamed Then
        RaiseEvent FolderRenamed(Item1, Item2)
      End If
  End Select
End Sub

Private Sub Class_Initialize()
  RegisterShellChangeClass Me
End Sub

Private Sub Class_Terminate()
  If IsReceiving Then
    StopReceiving
  End If
  UnregisterShellChangeClass Me
End Sub
