VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WavePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function LengthCall Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As MCI_STATUS_PARMS) As Long
Private Declare Function PositionCall1 Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As MCI_STATUS_PARMS) As Long
Private Declare Function PositionCall2 Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As MCI_SEEK_PARMS) As Long
Private Declare Function OpenCall1 Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As MCI_OPEN_PARMS) As Long
Private Declare Function OpenCall2 Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As MCI_SET_PARMS) As Long
Private Declare Function CloseCall Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As MCI_GENERIC_PARMS) As Long
Private Declare Function PlayCall Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As MCI_PLAY_PARMS) As Long
Private Declare Function OtherCall Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As Long) As Long

Dim m_FileName As String

Dim DeviceID As Long

Public Property Get FileName() As String
  FileName = m_FileName
End Property
Public Property Get Length() As Long
  Dim cParms As MCI_STATUS_PARMS
  If IsFileOpen Then
    cParms.dwItem = MCI_STATUS_LENGTH
    LengthCall DeviceID, MCI_STATUS, MCI_STATUS_ITEM, cParms
    Length = cParms.dwReturn
  End If
End Property
Public Property Get Position() As Long
  Dim cParms As MCI_STATUS_PARMS
  Dim cResult As Long
  If IsFileOpen Then
    cParms.dwItem = MCI_STATUS_POSITION
    cResult = PositionCall1(DeviceID, MCI_STATUS, MCI_WAIT Or MCI_STATUS_ITEM, cParms)
    If cResult = 0 Then
      If (cParms.dwReturn < 0 Or cParms.dwReturn > Length) Then
        Position = 0
      Else
        Position = cParms.dwReturn
      End If
    Else
      Position = 0
    End If
  End If
End Property
Public Property Let FileName(New_FileName As String)
  If Not LCase(m_FileName) = LCase(New_FileName) Then
    m_FileName = New_FileName
  End If
End Property
Public Property Let Position(New_Position As Long)
  Dim cParms As MCI_SEEK_PARMS
  If IsFileOpen Then
    cParms.dwTo = New_Position
    PositionCall2 DeviceID, MCI_SEEK, MCI_TO, cParms
  End If
End Property
Public Property Let Command(New_Command As MULTIMEDIA_COMMAND)
  Select Case New_Command
    Case MULTIMEDIA_COMMAND.mcOpen
      OpenFile
    Case MULTIMEDIA_COMMAND.mcClose
      CloseFile
    Case MULTIMEDIA_COMMAND.mcPlay
      StartPlay
    Case MULTIMEDIA_COMMAND.mcPause
      PausePlay
    Case MULTIMEDIA_COMMAND.mcStop
      StopPlay
    Case MULTIMEDIA_COMMAND.mcRewind
      RewindPlayer
  End Select
End Property

Private Function IsFileOpen() As Boolean
  IsFileOpen = Not (DeviceID = 0)
End Function

Private Sub CloseFile()
  Dim cParms As MCI_GENERIC_PARMS
  If IsFileOpen Then
    CloseCall DeviceID, MCI_CLOSE, MCI_WAIT, cParms
    DeviceID = 0
  End If
End Sub

Private Sub OpenFile()
  Dim cParms As MCI_SET_PARMS
  Dim cResult As Long
  Dim cError As String
  Dim cOpenParms As MCI_OPEN_PARMS
  CloseFile
  If FileExistsA(m_FileName) Then
    With cOpenParms
      .dwCallback = 0
      .wDeviceID = 0
      .lpstrDeviceType = vbNullString
      .lpstrAlias = vbNullString
      .lpstrElementName = m_FileName
    End With
    cResult = OpenCall1(0, MCI_OPEN, MCI_WAIT Or MCI_OPEN_ELEMENT, cOpenParms)
    DeviceID = cOpenParms.wDeviceID
    If Not cResult = 0 Then
      Err.Raise cResult
    Else
      cParms.dwTimeFormat = MCI_FORMAT_MILLISECONDS
      cResult = OpenCall2(DeviceID, MCI_SET, MCI_WAIT Or MCI_SET_TIME_FORMAT, cParms)
      If Not cResult = 0 Then
        Err.Raise cResult
      End If
    End If
  End If
End Sub

Private Sub StartPlay()
  Dim cParms As MCI_PLAY_PARMS
  If IsFileOpen Then
    PlayCall DeviceID, MCI_PLAY, 0, cParms
  Else
    OpenFile
    If IsFileOpen Then
      StartPlay
    End If
  End If
End Sub

Private Sub PausePlay()
  If IsFileOpen Then
    OtherCall DeviceID, MCI_STOP, MCI_WAIT, 0
  End If
End Sub

Private Sub StopPlay()
  If IsFileOpen Then
    OtherCall DeviceID, MCI_STOP, MCI_WAIT, 0
    RewindPlayer
  End If
End Sub

Private Sub RewindPlayer()
  If IsFileOpen Then
    OtherCall DeviceID, MCI_SEEK, MCI_WAIT Or MCI_SEEK_TO_START, 0
    Position = 0
  End If
End Sub

Private Sub Class_Terminate()
  StopPlay
  CloseFile
End Sub
