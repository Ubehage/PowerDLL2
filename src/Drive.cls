VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_DriveLetter As String
Dim m_DriveType As DRIVETYPE_Constants

Dim DiskInfo As DRIVEINFO

Public Property Get DriveLetter() As String
  DriveLetter = m_DriveLetter
End Property
Public Property Get DriveType() As DRIVETYPE_Constants
  DriveType = m_DriveType
End Property
Public Property Get DriveTypeName() As String
  DriveTypeName = GetDriveTypeName(, Me)
End Property
Public Property Get TotalSpace() As Double
  UpdateDiskSpaceInfo
  TotalSpace = DiskInfo.drvSpaceTotal
End Property
Public Property Get UsedSpace() As Double
  UpdateDiskSpaceInfo
  UsedSpace = DiskInfo.drvSpaceUsed
End Property
Public Property Get FreeSpace()
  UpdateDiskSpaceInfo
  FreeSpace = DiskInfo.drvSpaceFree
End Property
Public Property Get VolumeLabel() As String
  UpdateDiskVolumeInfo
  VolumeLabel = DiskInfo.drvVolumeName
End Property
Public Property Get SerialNumber() As String
  UpdateDiskVolumeInfo
  SerialNumber = Trim(DiskInfo.drvSerialNo)
End Property
Public Property Get FileSystem() As String
  UpdateDiskVolumeInfo
  FileSystem = DiskInfo.drvFileSystemName
End Property
Public Property Get MaxCharactersInPath() As Long
  UpdateDiskVolumeInfo
  MaxCharactersInPath = DiskInfo.drvFileSystemSupport
End Property
Public Property Get Icon() As IPictureDisp
  Set Icon = GetAssociatedIconA((m_DriveLetter + ":\"), True)
End Property
Friend Property Let DriveLetter(New_DriveLetter As String)
  m_DriveLetter = New_DriveLetter
  m_DriveType = GetDriveTypeA(m_DriveLetter)
End Property
Public Property Let VolumeLabel(New_VolumeLabel As String)
  SetVolumeLabel (m_DriveLetter + ":\"), New_VolumeLabel
End Property

Private Sub UpdateDiskSpaceInfo()
  GetDiskSpaceInformation m_DriveLetter, DiskInfo
End Sub

Private Sub UpdateDiskVolumeInfo()
  GetDiskVolumeInformation m_DriveLetter, DiskInfo
End Sub
