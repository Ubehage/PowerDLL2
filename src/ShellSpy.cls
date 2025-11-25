VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim CanEnable As Boolean

Dim m_FolderEvents As Boolean
Dim m_FileEvents As Boolean
Dim m_DriveEvents As Boolean
Dim m_MediaEvents As Boolean
Dim m_NetworkEvents As Boolean
Dim m_Enabled As Boolean

Public Event FolderCreated(Path As String, DisplayName As String, PIDL As Long)
Public Event FolderRenamed(Path As String, Path2 As String, DisplayName As String, DisplayName2 As String, PIDL As Long, PIDL2 As Long)
Public Event FolderDeleted(Path As String, PIDL As Long)
Public Event FolderUpdated(Path As String, DisplayName As String, PIDL As Long)
Public Event FileCreated(Path As String, DisplayName As String, PIDL As Long)
Public Event FileRenamed(Path As String, Path2 As String, DisplayName As String, DisplayName2 As String, PIDL As Long, PIDL2 As Long)
Public Event FileDeleted(Path As String, PIDL As Long)
Public Event FileUpdated(Path As String, DisplayName As String, PIDL As Long)
Public Event DriveAdded(Path As String, PIDL As Long)
Public Event DriveRemoved(Path As String, PIDL As Long)
Public Event DriveSpaceChanged(Path As String, PIDL As Long)
Public Event DriveUpdated(Path As String, PIDL As Long)
Public Event MediaInserted(Path As String, PIDL As Long)
Public Event MediaRemoved(Path As String, PIDL As Long)
Public Event NetworkShareAdded(Path As String, DisplayName As String, PIDL As Long)
Public Event NetworkShareRemoved(Path As String, PIDL As Long)
Public Event SystemImageChanged()

Public Property Get FolderEvents() As Boolean
  FolderEvents = m_FolderEvents
End Property
Public Property Get FileEvents() As Boolean
  FileEvents = m_FileEvents
End Property
Public Property Get DriveEvents() As Boolean
  DriveEvents = m_DriveEvents
End Property
Public Property Get MediaEvents() As Boolean
  MediaEvents = m_MediaEvents
End Property
Public Property Get NetworkEvents() As Boolean
  NetworkEvents = m_NetworkEvents
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Let FolderEvents(New_FolderEvents As Boolean)
  If Not m_FolderEvents = New_FolderEvents Then
    m_FolderEvents = New_FolderEvents
  End If
End Property
Public Property Let FileEvents(New_FileEvents As Boolean)
  If Not m_FileEvents = New_FileEvents Then
    m_FileEvents = New_FileEvents
  End If
End Property
Public Property Let DriveEvents(New_DriveEvents As Boolean)
  If Not m_DriveEvents = New_DriveEvents Then
    m_DriveEvents = New_DriveEvents
  End If
End Property
Public Property Let MediaEvents(New_MediaEvents As Boolean)
  If Not m_MediaEvents = New_MediaEvents Then
    m_MediaEvents = New_MediaEvents
  End If
End Property
Public Property Let NetworkEvents(New_NetworkEvents As Boolean)
  If Not m_NetworkEvents = New_NetworkEvents Then
    m_NetworkEvents = New_NetworkEvents
  End If
End Property
Public Property Let Enabled(New_Enabled As Boolean)
  If Not m_Enabled = New_Enabled Then
    If CanEnable Then
      m_Enabled = New_Enabled
    End If
  End If
End Property

Friend Sub EventCaller(EventID As mbShellEventConstants, DisplayName As String, Path As String, PIDL As Long, DisplayName2 As String, Path2 As String, PIDL2 As Long)
  Select Case EventID
    Case MBShellSpy.mbShellEventConstants.mbDriveAdded
      RaiseEvent DriveAdded(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbDriveFreeSpaceChanged
      RaiseEvent DriveSpaceChanged(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbDriveIconAdded
      RaiseEvent DriveUpdated(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbDriveRemoved
      RaiseEvent DriveRemoved(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbFolderCreated
      RaiseEvent FolderCreated(Path, DisplayName, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbFolderDeleted
      RaiseEvent FolderDeleted(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbFolderRenamed
      RaiseEvent FolderRenamed(Path, Path2, DisplayName, DisplayName2, PIDL, PIDL2)
    Case MBShellSpy.mbShellEventConstants.mbFolderUpdated
      RaiseEvent FolderUpdated(Path, DisplayName, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbItemCreated
      RaiseEvent FileCreated(Path, DisplayName, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbItemDeleted
      RaiseEvent FileDeleted(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbItemRenamed
      RaiseEvent FileRenamed(Path, Path2, DisplayName, DisplayName2, PIDL, PIDL2)
    Case MBShellSpy.mbShellEventConstants.mbItemUpdated
      RaiseEvent FileUpdated(Path, DisplayName, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbMediaInserted
      RaiseEvent MediaInserted(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbMediaRemoved
      RaiseEvent MediaRemoved(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbNetworkShareAdded
      RaiseEvent NetworkShareAdded(Path, DisplayName, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbNetworkShareRemoved
      RaiseEvent NetworkShareRemoved(Path, PIDL)
    Case MBShellSpy.mbShellEventConstants.mbSysImageChanged
      RaiseEvent SystemImageChanged
  End Select
End Sub

Private Sub Class_Initialize()
  If RegisterShellSpyClass(Me) Then
    CanEnable = True
  End If
End Sub

Private Sub Class_Terminate()
  If CanEnable Then
    UnRegisterShellSpyClass Me
  End If
End Sub
