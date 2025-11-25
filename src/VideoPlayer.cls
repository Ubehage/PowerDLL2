VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_FileName As String
Dim m_hWndParent As Long
Dim m_Stretch As Boolean
Dim m_VideoWidth As Long
Dim m_VideoHeight As Long
Dim m_Left As Long
Dim m_Top As Long
Dim m_Width As Long
Dim m_Height As Long
Dim m_hWndVideo As Long

Dim DeviceID As Long

Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As AVIRECT) As Long

Public Property Get FileName() As String
  FileName = m_FileName
End Property
Public Property Get hWndParent() As Long
  hWndParent = m_hWndParent
End Property
Public Property Get Stretch() As Boolean
  Stretch = m_Stretch
End Property
Public Property Get VideoSize() As VideoSize
  Dim vRect As AVIRECT
  If IsFileOpen Then
    If Not m_hWndVideo = 0 Then
      GetClientRect m_hWndVideo, vRect
      With VideoSize
        .Width = vRect.Right
        .Height = vRect.Left
      End With
    End If
  End If
End Property
Public Property Get Left() As Long
  Left = m_Left
End Property
Public Property Get Top() As Long
  Top = m_Top
End Property
Public Property Get Width() As Long
  Width = m_Width
End Property
Public Property Get Height() As Long
  Height = m_Height
End Property
Public Property Get hWndVideo() As Long
  hWndVideo = m_hWndVideo
End Property
Public Property Get Length() As Long
  Dim cParms As MCI_STATUS_PARMS
  If IsFileOpen Then
    cParms.dwItem = MCI_STATUS_LENGTH
    mciSendCommand DeviceID, MCI_STATUS, MCI_STATUS_ITEM, cParms
    Length = cParms.dwReturn
  End If
End Property
Public Property Get Position() As Long
  Dim cParms As MCI_STATUS_PARMS
  If IsFileOpen Then
    cParms.dwItem = MCI_STATUS_POSITION
    If mciSendCommand(DeviceID, MCI_STATUS, MCI_STATUS_ITEM, cParms) = 0 Then
      If (cParms.dwReturn >= 0 And cParms.dwReturn <= Length) Then
        Position = cParms.dwReturn
      End If
    End If
  End If
End Property


Public Property Let FileName(New_FileName As String)
  If Not LCase(m_FileName) = LCase(New_FileName) Then
    m_FileName = New_FileName
  End If
End Property
Public Property Let hWndParent(New_hWndParent As Long)
  If Not m_hWndParent = New_hWndParent Then
    m_hWndParent = New_hWndParent
  End If
End Property
Public Property Let Stretch(New_Stretch As Boolean)
  If Not m_Stretch = New_Stretch Then
    m_Stretch = New_Stretch
  End If
End Property
Public Property Let Left(New_Left As Long)
  If Not m_Left = New_Left Then
    m_Left = New_Left
  End If
End Property
Public Property Let Top(New_Top As Long)
  If Not m_Top = New_Top Then
    m_Top = New_Top
  End If
End Property
Public Property Let Width(New_Width As Long)
  If Not m_Width = New_Width Then
    m_Width = New_Width
  End If
End Property
Public Property Let Height(New_Height As Long)
  If Not m_Height = New_Height Then
    m_Height = New_Height
  End If
End Property
Public Property Let Position(New_Position As Long)
  Dim cParms As MCI_SEEK_PARMS
  If IsFileOpen Then
    cParms.dwTo = New_Position
    mciSendCommand DeviceID, MCI_SEEK, MCI_TO, cParms
  End If
End Property
Public Property Let Command(New_Command As MULTIMEDIA_COMMAND)
  Select Case New_Command
    Case MULTIMEDIA_COMMAND.mcOpen
      OpenFile
    Case MULTIMEDIA_COMMAND.mcClose
      CloseFile
    Case MULTIMEDIA_COMMAND.mcPlay
      PlayVideo
    Case MULTIMEDIA_COMMAND.mcPause
      PauseVideo
    Case MULTIMEDIA_COMMAND.mcStop
      StopVideo
    Case MULTIMEDIA_COMMAND.mcRewind
      RewindVideo
  End Select
End Property

Public Sub ResizeVideo()
  MoveVideoWindow
End Sub

Private Function IsFileOpen() As Boolean
  IsFileOpen = Not (DeviceID = 0)
End Function

Private Sub OpenFile()
  Dim cOpenParms As MCI_DGV_OPEN_PARMS
  Dim cWindowParms As MCI_DGV_WINDOW_PARMS
  Dim cStatusParms As MCI_DGV_STATUS_PARMS
  Dim cResult As Long
  CloseFile
  If FileExistsA(m_FileName) Then
    If Not m_hWndParent = 0 Then
      With cOpenParms
        .dwCallback = 0
        .wDeviceID = 0
        .lpstrDeviceType = 0
        .lpstrElementName = m_FileName
        .lpstrAlias = vbNullString
        .dwStyle = WS_CHILD
        .hWndParent = m_hWndParent
      End With
      cResult = mciSendCommand(0, MCI_OPEN, MCI_OPEN_ELEMENT Or MCI_DGV_OPEN_PARENT Or MCI_DGV_OPEN_WS, cOpenParms)
      If cResult = 0 Then
        DeviceID = cOpenParms.wDeviceID
        With cWindowParms
          .dwCallback = 0
          .hWnd = 0
          .nCmdShow = SW_SHOW
          .lpstrText = vbNullString
        End With
        mciSendCommand DeviceID, MCI_WINDOW, MCI_DGV_WINDOW_STATE, cWindowParms
        cStatusParms.dwItem = MCI_DGV_STATUS_HWND
        mciSendCommand DeviceID, MCI_STATUS, MCI_STATUS_ITEM, cStatusParms
        m_hWndVideo = cStatusParms.dwReturn
      Else
        Err.Raise cResult
      End If
    End If
  End If
End Sub

Private Sub CloseFile()
  Dim cParms As MCI_GENERIC_PARMS
  StopVideo
  If IsFileOpen Then
    mciSendCommand DeviceID, MCI_CLOSE, 0, cParms
    DeviceID = 0
  End If
End Sub

Private Sub PlayVideo()
  Dim cParms As MCI_PLAY_PARMS
  If IsFileOpen Then
    mciSendCommand DeviceID, MCI_PLAY, MCI_MCIAVI_PLAY_WINDOW, cParms
  Else
    OpenFile
    If IsFileOpen Then
      PlayVideo
    End If
  End If
End Sub

Private Sub PauseVideo()
  Dim cParms As MCI_GENERIC_PARMS
  If IsFileOpen Then
    mciSendCommand DeviceID, MCI_PAUSE, 0, cParms
  End If
End Sub

Private Sub StopVideo()
  If IsFileOpen Then
    mciSendCommand DeviceID, MCI_STOP, 0, 0
    RewindVideo
  End If
End Sub

Private Sub RewindVideo()
  If IsFileOpen Then
    mciSendCommand DeviceID, MCI_SEEK, MCI_WAIT Or MCI_SEEK_TO_START, 0
    Position = 0
  End If
End Sub

Private Sub MoveVideoWindow()
  Dim vRect As AVIRECT
  Dim vParms As MCI_DGV_RECT_PARMS
  If IsFileOpen Then
    If Not m_hWndVideo = 0 Then
      If m_Stretch Then
        MoveWindow m_hWndVideo, m_Left, m_Top, m_Width, m_Height, 1
      Else
        mciSendCommand DeviceID, MCI_WHERE, MCI_DGV_WHERE_SOURCE, vParms
        vRect = vParms.rc
        vRect.Left = m_Left
        vRect.Top = m_Top
        MoveWindow m_hWndVideo, vRect.Left, vRect.Top, vRect.Right, vRect.Bottom, 1
      End If
    End If
  End If
End Sub

Private Sub Class_Terminate()
  CloseFile
End Sub
