Attribute VB_Name = "modGlobal"
Option Explicit

Public Const MAX_LENGTH = 260
Public Const MAX_PATH = 260

Public Const S_OK = 0
Public Const S_FALSE = 1

Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_OPENICON = &H2
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_PIDL = &H8
Public Const SHGFI_USEFILEATTRIBUTES = &H10
Public Const SHGFI_ICON = &H100
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_ATTRIBUTES = &H800
Public Const SHGFI_ICONLOCATION = &H1000
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_LINKOVERLAY = &H8000
Public Const SHGFI_SELECTED = &H10000
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Const SHGFP_TYPE_CURRENT = &H0
Public Const SHGFP_TYPE_DEFAULT = &H1

Public Const ILD_TRANSPARENT = &H1

Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32

Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000

Public Const CDS_FORCE = &H80000000

Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const OPEN_EXISTING As Long = 3

Public Const IOCTL_STORAGE_GET_MEDIA_TYPES_EX As Long = &H2D0C04

Public Const FILE_DEVICE_CD_ROM = &H2
Public Const FILE_DEVICE_DVD = &H33

Public Const MAX_COMPUTERNAME As Long = 16
Public Const REG_BINARY As Long = &H3
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const STANDARD_RIGHTS_READ As Long = &H20000
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY As Long = &H10
Public Const SYNCHRONIZE As Long = &H100000
Public Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Type DEVICE_MEDIA_INFO
 Cylinders As Double
 MediaType As Long
 TracksPerCylinder As Long
 SectorsPerTrack As Long
 BytesPerSector As Long
 NumberMediaSides As Long
 MediaCharacteristics As Long
End Type

Public Type GET_MEDIA_TYPES
 DeviceType As Long
 MediaInfoCount As Long
 MediaInfo(10) As DEVICE_MEDIA_INFO
End Type

Public Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Public Type OSVERSIONINFO
  OSVSize As Long
  dwVerMajor As Long
  dwVerMinor As Long
  dwBuildNumber As Long
  PlatformID As Long
  szCSDVersion As String * 128
End Type

Public Type SHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Public Type DRIVEINFO
  drvSpaceFree As Currency
  drvSpaceFreeToCaller As Currency
  drvSpaceUsed As Currency
  drvSpaceTotal As Currency
  drvVolumeName As String
  drvSerialNo As String
  drvFileSystemName As String
  drvFileSystemSupport As Long
  drvFileSystemFlags As Long
End Type

Public Type MCI_OPEN_PARMS
  dwCallback As Long
  wDeviceID As Long
  lpstrDeviceType As String
  lpstrElementName As String
  lpstrAlias As String
End Type

Public Type MCI_GENERIC_PARMS
  dwCallback As Long
End Type

Public Type MCI_SET_PARMS
  dwCallback As Long
  dwTimeFormat As Long
  dwAudio As Long
End Type

Public Type MCI_PLAY_PARMS
  dwCallback As Long
  dwFrom As Long
  dwTo As Long
End Type

Public Type MCI_STATUS_PARMS
  dwCallback As Long
  dwReturn As Long
  dwItem As Long
  dwTrack As Integer
End Type

Public Type MCI_SEEK_PARMS
  dwCallback As Long
  dwTo As Long
End Type

Public Type MCI_RECORD_PARMS
  dwCallback As Long
  dwFrom As Long
  dwTo As Long
End Type

Public Type MCI_SAVE_PARMS
  dwCallback As Long
  lpFileName As String
End Type

Public Type AVIRECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type MCI_DGV_RECT_PARMS
  dwCallback As Long
  rc As AVIRECT
End Type

Public Type MCI_DGV_OPEN_PARMS
  dwCallback As Long
  wDeviceID As Long
  lpstrDeviceType As String
  lpstrElementName As String
  lpstrAlias As String
  dwStyle As Long
  hWndParent As Long
End Type

Public Type MCI_DGV_WINDOW_PARMS
  dwCallback As Long
  hWnd As Long
  nCmdShow As Long
  lpstrText As String
End Type

Public Type MCI_DGV_STATUS_PARMS
  dwCallback As Long
  dwReturn As Long
  dwItem As Long
  dwTrack As Long
  lpstrDrive As String
  dwReference As Long
End Type

Public Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

Public Type TIME_ZONE_INFORMATION
  Bias As Long
  StandardName(0 To 63) As Byte
  StandardDate As SYSTEMTIME
  StandardBias As Long
  DaylightName(0 To 63) As Byte
  DaylightDate As SYSTEMTIME
  DaylightBias As Long
End Type

Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZone As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileDate As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Public Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long

Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long

Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long

Public Declare Function MBRestartSystemDialog Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function MBShutdownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (byvalhDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long
Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
'Public Declare Function SHGetFileInfoPidl Lib "shell32" Alias "SHGetFileInfoA" (ByVal pidl As Long, ByVal dwFileAttributes As Long, psfib As SHFILEINFOBYTE, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function SHFileOperation Lib "shell32" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHAutoComplete Lib "Shlwapi.dll" (ByVal hWndEdit As Long, ByVal dwFlags As Long) As Long

Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
  
Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpcc As CHOOSECOLORSTRUCT) As Long

Public Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long

Public Declare Function ExtractIconB Lib "shell32" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Declare Function CreateProcessA Lib "kernel32" (ByVal lpAppName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliSeconds As Long) As Long

Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal sResult As String) As Long

'Public Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" (ByVal hWnd As Long, ByVal uFlags As Long, ByVal dwEventID As Long, ByVal uMsg As Long, ByVal cItems As Long, lpps As PIDLSTRUCT) As Long
'Public Declare Function SHChangeNotifyDeregister Lib "shell32" Alias "#4" (ByVal hNotify As Long) As Boolean
'Public Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDrive As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function InetIsOffline Lib "url.dll" (ByVal dwFlags As Long) As Long

Public Function TrimNull(TrimString As String) As String
  TrimNull = Left(TrimString, lstrlenW(StrPtr(TrimString)))
End Function

Public Sub UpdateDiskInformation(Drive As String, DiskInfo As DRIVEINFO)
  GetDiskSpaceInformation Drive, DiskInfo
  GetDiskVolumeInformation Drive, DiskInfo
End Sub

Public Sub GetDiskSpaceInformation(Drive As String, DiskInfo As DRIVEINFO)
  Dim BytesFreeToCaller As Currency
  Dim TotalBytes As Currency
  Dim TotalFreeBytes As Currency
  If GetDiskFreeSpaceEx((Left(Drive, 1) + ":"), BytesFreeToCaller, TotalBytes, TotalFreeBytes) = 1 Then
    With DiskInfo
      .drvSpaceTotal = (TotalBytes * 10000)
      .drvSpaceFree = (TotalFreeBytes * 10000)
      .drvSpaceFreeToCaller = (BytesFreeToCaller * 10000)
      .drvSpaceUsed = ((TotalBytes - TotalFreeBytes) * 10000)
    End With
  End If
End Sub

Public Sub GetDiskVolumeInformation(Drive As String, DiskInfo As DRIVEINFO)
  Dim pos As Integer
  Dim HiWord As Long
  Dim HiHexStr As String
  Dim LoWord As Long
  Dim LoHexStr As String
  Dim VolumeSN As Long
  Dim MaxFNLen As Long
  Dim vnSize As Long
  Dim fnSize As Long
  With DiskInfo
    .drvVolumeName = Space(14)
    .drvFileSystemName = Space(32)
    vnSize = Len(.drvVolumeName)
    fnSize = Len(.drvFileSystemName)
    If GetVolumeInformation((Left(Drive, 1) + ":\"), .drvVolumeName, vnSize, VolumeSN, MaxFNLen, .drvFileSystemFlags, .drvFileSystemName, fnSize) Then
      pos = InStr(.drvVolumeName, Chr(0))
      If pos Then
        .drvVolumeName = Left(.drvVolumeName, (pos - 1))
      End If
      If Len(Trim(.drvVolumeName)) = 0 Then
        .drvVolumeName = GetDriveTypeName(Drive)
      End If
      pos = InStr(.drvFileSystemName, Chr(0))
      If pos Then
        .drvFileSystemName = Left(.drvFileSystemName, (pos - 1))
      End If
      .drvFileSystemSupport = MaxFNLen
      .drvSerialNo = Hex(VolumeSN)
    Else
      .drvVolumeName = GetDriveTypeName(Drive)
    End If
  End With
End Sub

Public Function GetDriveTypeName(Optional Drive As String, Optional DriveObject As PowerDLL2.Drive) As String
  Dim dType As DRIVETYPE_Constants
  If Not DriveObject Is Nothing Then
    dType = DriveObject.DriveType
  ElseIf Not Drive = "" Then
    dType = GetDriveObjectFromDriveLetterA(Drive).DriveType
  Else
    Exit Function
  End If
  Select Case dType
    Case dtRemoveable
      GetDriveTypeName = (("3½" + Chr(34)) + " Floppy")
    Case dtHardDisk
      GetDriveTypeName = "Local Disk"
    Case dtCD
      GetDriveTypeName = "CD"
    Case dtDVD
      GetDriveTypeName = "DVD"
    Case dtNetwork
      GetDriveTypeName = "Network Disk"
    Case dtRamDrive
      GetDriveTypeName = "RAM Disk"
  End Select
End Function

Public Function GetExplorerFile() As String
  GetExplorerFile = (FixPathA(GetSpecialFolderPathA(sfWindowsDirectory)) + "\explorer.exe")
End Function

Public Function GetStartFile() As String
  GetStartFile = "start"
End Function

Public Function GetShell32File() As String
  GetShell32File = (FixPathA(GetSpecialFolderPathA(sfSystemDirectory)) + "\shell32.dll")
End Function

Public Function SetMaxDecimals(Value As Double, Decimals As Long) As String
  Dim dValue As String
  Dim pos As Long
  dValue = Trim(Str(Value))
  pos = InStr(dValue, ".")
  If Not pos = 0 Then
    If (Len(dValue) - pos) > Decimals Then
      dValue = Left(dValue, (pos + Decimals))
    End If
  End If
  SetMaxDecimals = dValue
End Function

Public Function Replace2(Text As String, Find As String, ReplaceWith As String) As String
  Dim i As Long
  Dim nText As String
  nText = Text
  Do Until i >= (Len(nText) - (Len(Find) - 1))
    i = (i + 1)
    If LCase(Mid(nText, i, Len(Find))) = LCase(Find) Then
      nText = ((Left(nText, (i - 1)) + ReplaceWith) + Right(nText, (Len(nText) - (i + (Len(Find) - 1)))))
      i = (i + (Len(ReplaceWith) - 1))
    End If
  Loop
  Replace2 = nText
End Function

Public Function GetKeyConstantFromCode(KeyCode As Integer) As VBRUN.KeyCodeConstants
  Select Case KeyCode
    Case vbKey0
      GetKeyConstantFromCode = vbKey0
    Case vbKey1
      GetKeyConstantFromCode = vbKey1
    Case vbKey2
      GetKeyConstantFromCode = vbKey2
    Case vbKey3
      GetKeyConstantFromCode = vbKey3
    Case vbKey4
      GetKeyConstantFromCode = vbKey4
    Case vbKey5
      GetKeyConstantFromCode = vbKey5
    Case vbKey6
      GetKeyConstantFromCode = vbKey6
    Case vbKey7
      GetKeyConstantFromCode = vbKey7
    Case vbKey8
      GetKeyConstantFromCode = vbKey8
    Case vbKey9
      GetKeyConstantFromCode = vbKey9
    Case vbKeyA
      GetKeyConstantFromCode = vbKeyA
    Case vbKeyB
      GetKeyConstantFromCode = vbKeyB
    Case vbKeyC
      GetKeyConstantFromCode = vbKeyC
    Case vbKeyD
      GetKeyConstantFromCode = vbKeyD
    Case vbKeyE
      GetKeyConstantFromCode = vbKeyE
    Case vbKeyF
      GetKeyConstantFromCode = vbKeyF
    Case vbKeyG
      GetKeyConstantFromCode = vbKeyG
    Case vbKeyH
      GetKeyConstantFromCode = vbKeyH
    Case vbKeyI
      GetKeyConstantFromCode = vbKeyI
    Case vbKeyJ
      GetKeyConstantFromCode = vbKeyJ
    Case vbKeyK
      GetKeyConstantFromCode = vbKeyK
    Case vbKeyL
      GetKeyConstantFromCode = vbKeyL
    Case vbKeyM
      GetKeyConstantFromCode = vbKeyM
    Case vbKeyN
      GetKeyConstantFromCode = vbKeyN
    Case vbKeyO
      GetKeyConstantFromCode = vbKeyO
    Case vbKeyP
      GetKeyConstantFromCode = vbKeyP
    Case vbKeyQ
      GetKeyConstantFromCode = vbKeyQ
    Case vbKeyR
      GetKeyConstantFromCode = vbKeyR
    Case vbKeyS
      GetKeyConstantFromCode = vbKeyS
    Case vbKeyT
      GetKeyConstantFromCode = vbKeyT
    Case vbKeyU
      GetKeyConstantFromCode = vbKeyU
    Case vbKeyV
      GetKeyConstantFromCode = vbKeyV
    Case vbKeyW
      GetKeyConstantFromCode = vbKeyW
    Case vbKeyX
      GetKeyConstantFromCode = vbKeyX
    Case vbKeyY
      GetKeyConstantFromCode = vbKeyY
    Case vbKeyZ
      GetKeyConstantFromCode = vbKeyZ
    Case vbKeyF1
      GetKeyConstantFromCode = vbKeyF1
    Case vbKeyF2
      GetKeyConstantFromCode = vbKeyF2
    Case vbKeyF3
      GetKeyConstantFromCode = vbKeyF3
    Case vbKeyF4
      GetKeyConstantFromCode = vbKeyF4
    Case vbKeyF5
      GetKeyConstantFromCode = vbKeyF5
    Case vbKeyF6
      GetKeyConstantFromCode = vbKeyF6
    Case vbKeyF7
      GetKeyConstantFromCode = vbKeyF7
    Case vbKeyF8
      GetKeyConstantFromCode = vbKeyF8
    Case vbKeyF9
      GetKeyConstantFromCode = vbKeyF9
    Case vbKeyF10
      GetKeyConstantFromCode = vbKeyF10
    Case vbKeyF11
      GetKeyConstantFromCode = vbKeyF11
    Case vbKeyF12
      GetKeyConstantFromCode = vbKeyF12
    Case vbKeyAdd
      GetKeyConstantFromCode = vbKeyAdd
    Case vbKeyBack
      GetKeyConstantFromCode = vbKeyBack
    Case vbKeyCancel
      GetKeyConstantFromCode = vbKeyCancel
    Case vbKeyCapital
      GetKeyConstantFromCode = vbKeyCapital
    Case vbKeyClear
      GetKeyConstantFromCode = vbKeyClear
    Case vbKeyControl
      GetKeyConstantFromCode = vbKeyControl
    Case vbKeyDecimal
      GetKeyConstantFromCode = vbKeyDecimal
    Case vbKeyDelete
      GetKeyConstantFromCode = vbKeyDelete
    Case vbKeyDivide
      GetKeyConstantFromCode = vbKeyDivide
    Case vbKeyDown
      GetKeyConstantFromCode = vbKeyDown
    Case vbKeyEnd
      GetKeyConstantFromCode = vbKeyEnd
    Case vbKeyEscape
      GetKeyConstantFromCode = vbKeyEscape
    Case vbKeyExecute
      GetKeyConstantFromCode = vbKeyExecute
    Case vbKeyHelp
      GetKeyConstantFromCode = vbKeyHelp
    Case vbKeyHome
      GetKeyConstantFromCode = vbKeyHome
    Case vbKeyInsert
      GetKeyConstantFromCode = vbKeyInsert
    Case vbKeyLButton
      GetKeyConstantFromCode = vbKeyLButton
    Case vbKeyLeft
      GetKeyConstantFromCode = vbKeyLeft
    Case vbKeyMButton
      GetKeyConstantFromCode = vbKeyMButton
    Case vbKeyMenu
      GetKeyConstantFromCode = vbKeyMenu
    Case vbKeyMultiply
      GetKeyConstantFromCode = vbKeyMultiply
    Case vbKeyNumlock
      GetKeyConstantFromCode = vbKeyNumlock
    Case vbKeyNumpad0
      GetKeyConstantFromCode = vbKeyNumpad0
    Case vbKeyNumpad1
      GetKeyConstantFromCode = vbKeyNumpad1
    Case vbKeyNumpad2
      GetKeyConstantFromCode = vbKeyNumpad2
    Case vbKeyNumpad3
      GetKeyConstantFromCode = vbKeyNumpad3
    Case vbKeyNumpad4
      GetKeyConstantFromCode = vbKeyNumpad4
    Case vbKeyNumpad5
      GetKeyConstantFromCode = vbKeyNumpad5
    Case vbKeyNumpad6
      GetKeyConstantFromCode = vbKeyNumpad6
    Case vbKeyNumpad7
      GetKeyConstantFromCode = vbKeyNumpad7
    Case vbKeyNumpad8
      GetKeyConstantFromCode = vbKeyNumpad8
    Case vbKeyNumpad9
      GetKeyConstantFromCode = vbKeyNumpad9
    Case vbKeyPageDown
      GetKeyConstantFromCode = vbKeyPageDown
    Case vbKeyPageUp
      GetKeyConstantFromCode = vbKeyPageUp
    Case vbKeyPause
      GetKeyConstantFromCode = vbKeyPause
    Case vbKeyPrint
      GetKeyConstantFromCode = vbKeyPrint
    Case vbKeyRButton
      GetKeyConstantFromCode = vbKeyRButton
    Case vbKeyReturn
      GetKeyConstantFromCode = vbKeyReturn
    Case vbKeyRight
      GetKeyConstantFromCode = vbKeyRight
    Case vbKeyScrollLock
      GetKeyConstantFromCode = vbKeyScrollLock
    Case vbKeySelect
      GetKeyConstantFromCode = vbKeySelect
    Case vbKeySeparator
      GetKeyConstantFromCode = vbKeySeparator
    Case vbKeyShift
      GetKeyConstantFromCode = vbKeyShift
    Case vbKeySnapshot
      GetKeyConstantFromCode = vbKeySnapshot
    Case vbKeySpace
      GetKeyConstantFromCode = vbKeySpace
    Case vbKeySubtract
      GetKeyConstantFromCode = vbKeySubtract
    Case vbKeyTab
      GetKeyConstantFromCode = vbKeyTab
  End Select
End Function

Public Function GetShiftConstantFromCode(Shift As Integer) As VBRUN.ShiftConstants
  Dim sMask As VBRUN.ShiftConstants
  If (Shift And vbAltMask) Then
    sMask = (sMask Or vbAltMask)
  End If
  If (Shift And vbCtrlMask) Then
    sMask = (sMask Or vbCtrlMask)
  End If
  If (Shift And vbShiftMask) Then
    sMask = (sMask Or vbShiftMask)
  End If
  GetShiftConstantFromCode = sMask
End Function

Public Function PrepareForShutdown() As Boolean
  If IsWinNTA Then
    PrepareForShutdown = EnableShutdownPrivileges
  Else
    PrepareForShutdown = True
  End If
End Function

Public Function IsShellVersion(ByVal ShellVersion As Long) As Boolean
  Dim bSize As Long
  Dim nUnused As Long
  Dim lpBuffer As Long
  Dim nVerMajor As Integer
  Dim bBuffer() As Byte
  bSize = GetFileVersionInfoSize("shell32.dll", nUnused)
  If bSize > 0 Then
    ReDim bBuffer((bSize - 1)) As Byte
    GetFileVersionInfo "shell32.dll", 0&, bSize, bBuffer(0)
    If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
      CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
      IsShellVersion = (nVerMajor >= ShellVersion)
    End If
  End If
End Function

Public Function IsFullPath(PathString As String) As Boolean
  If Mid(PathString, 2, 1) = ":" Then
    If IsFolderA(PathString) Then
      IsFullPath = True
    ElseIf FileExistsA(PathString) Then
      IsFullPath = True
    End If
  ElseIf Left(PathString, 2) = "\\" Then
    If IsFolderA(PathString) Then
      IsFullPath = True
    ElseIf FileExistsA(PathString) Then
      IsFullPath = True
    End If
  End If
End Function

Public Function GetExitWindowsModeFromConstant(ExitMode As EXITWINDOWS_Constants) As Long
  Select Case ExitMode
    Case EXITWINDOWS_Constants.LogOff
      GetExitWindowsModeFromConstant = EWX_LOGOFF
    Case EXITWINDOWS_Constants.Restart
      GetExitWindowsModeFromConstant = EWX_REBOOT
    Case EXITWINDOWS_Constants.ShutDown
      GetExitWindowsModeFromConstant = EWX_SHUTDOWN
    Case Else
      GetExitWindowsModeFromConstant = EWX_NONE
  End Select
End Function

Public Function InvertColorValue(Value As Integer) As Integer
  Dim nValue As Integer
  nValue = (Value - 255)
  If nValue < 0 Then
    nValue = (-nValue)
  End If
  InvertColorValue = nValue
End Function

Public Function AddColorValues(Value1 As Integer, Value2 As Integer) As Integer
  If Value1 > Value2 Then
    AddColorValues = (Value1 - ((Value1 - Value2) / 2))
  ElseIf Value1 < Value2 Then
    AddColorValues = (Value1 + ((Value2 - Value1) / 2))
  Else
    AddColorValues = Value1
  End If
End Function

Public Function GetDateAndTimeFromFileTime(fTime As FILETIME) As DATE_AND_TIME
  Dim st As SYSTEMTIME
  Dim lt As SYSTEMTIME
  Dim tz As TIME_ZONE_INFORMATION
  If FileTimeToSystemTime(fTime, st) Then
    GetTimeZoneInformation tz
    SystemTimeToTzSpecificLocalTime tz, st, lt
    GetDateAndTimeFromFileTime.wDate = DateSerial(lt.wYear, lt.wMonth, lt.wDay)
    GetDateAndTimeFromFileTime.wTime = TimeSerial(lt.wHour, lt.wMinute, lt.wSecond)
  End If
End Function
