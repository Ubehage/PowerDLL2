VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type WINDOWS_VERSION_INFORMATION
  PlatformID As Long
  WindowsName As String
  VersionNumber As String
  ServicePack As String
  BuildNumber As String
End Type

Public Type DATE_AND_TIME
  wDate As Date
  wTime As Date
End Type

Public Enum DRIVETYPE_Constants
  dtRemoveable = 2
  dtHardDisk = 3
  dtNetwork = 4
  dtCD = 5
  dtRamDrive = 6
  dtDVD = 7
End Enum

Public Enum EXITWINDOWS_Constants
  LogOff = &H1
  Restart = &H2
  ShutDown = &H4
  Hibernate = &H8
  StandBy = &H10
End Enum

Public Enum ANIMATEWINDOW_Constants
  awHorizontalPositive = AW_HOR_POSITIVE
  awHorizontalNegative = AW_HOR_NEGATIVE
  awVerticalPositive = AW_VER_POSITIVE
  awVerticalNegative = AW_VER_NEGATIVE
  awCenter = AW_CENTER
  awHide = AW_HIDE
  awActivate = AW_ACTIVATE
  awSlide = AW_SLIDE
  awBlend = AW_BLEND
End Enum

Public Enum STARTMODE_Constants
  smSafeMode = 1
  smSafeModeWithNetwork = 2
  smNormalMode = 3
End Enum

Public Sub WindowOnTop(hWnd As Long, OnTop As Boolean)
  WindowOnTopA hWnd, OnTop
End Sub

Public Function GetDriveType(Drive As String) As DRIVETYPE_Constants
  GetDriveType = GetDriveTypeA(Drive)
End Function

Public Function IsProgramFile(File As String) As Boolean
  IsProgramFile = IsProgramFileA(File)
End Function

Public Function GetDriveLetters() As String
  GetDriveLetters = GetDriveLettersA
End Function

Public Function GetDriveObjectFromDriveLetter(Drive As String) As Drive
  Set GetDriveObjectFromDriveLetter = GetDriveObjectFromDriveLetterA(Drive)
End Function

Public Function GetWindowsVersion() As WINDOWS_VERSION_INFORMATION
  GetWindowsVersion = GetWindowsVersionA
End Function

Public Function IsWin95() As Boolean
  IsWin95 = IsWin95A
End Function

Public Function IsWin98() As Boolean
  IsWin98 = IsWin98A
End Function

Public Function IsWinME() As Boolean
  IsWinME = IsWinMEA
End Function

Public Function IsWinNT4() As Boolean
  IsWinNT4 = IsWinNT4A
End Function

Public Function IsWin2000() As Boolean
  IsWin2000 = IsWin2000A
End Function

Public Function IsWinXP() As Boolean
  IsWinXP = IsWinXPA
End Function

Public Function IsWinXPPlus() As Boolean
  IsWinXPPlus = IsWinXPPlusA
End Function

Public Function IsWinNT() As Boolean
  IsWinNT = IsWinNTA
End Function

Public Function GetWindowsName() As String
  GetWindowsName = GetWindowsNameA
End Function

Public Function ExitWindows(ExitMode As EXITWINDOWS_Constants, Optional ForceFlag As Boolean = False) As Boolean
  ExitWindows = ExitWindowsA(ExitMode, ForceFlag)
End Function

Public Function CanStandBy() As Boolean
  CanStandBy = CanStandByA
End Function

Public Function CanHibernate() As Boolean
  CanHibernate = CanHibernateA
End Function

Public Function GetMousePosition() As POINTAPI
  GetMousePosition = GetMousePositionA
End Function

Public Sub SetMousePosition(NewPosition As POINTAPI)
  SetMousePositionA NewPosition
End Sub

Public Sub ScreenToClient(hWnd As Long, Position As POINTAPI)
  ScreenToClientA hWnd, Position
End Sub

Public Sub ClientToScreen(hWnd As Long, Position As POINTAPI)
  ClientToScreenA hWnd, Position
End Sub

Public Sub StartCapture(hWnd As Long)
  StartCaptureA hWnd
End Sub

Public Sub EndCapture()
  EndCaptureA
End Sub

Public Function GetErrorDescription(ErrorCode As Long) As String
  GetErrorDescription = GetErrorDescriptionA(ErrorCode)
End Function

Public Sub SetDesktopWallpaper(WallpaperFile As String)
  SetDesktopWallpaperA WallpaperFile
End Sub

Public Sub ShowRestartDialog(Optional ExitMode As EXITWINDOWS_Constants = Restart, Optional hWnd As Long = 0&, Optional DialogMessage As String)
  ShowRestartDialogA ExitMode, hWnd, DialogMessage
End Sub

Public Sub ShuwShutdownDialog()
  ShowShutdownDialogA
End Sub

Public Sub ShowRunDialog(Optional hWnd As Long = 0&, Optional Title As String = vbNullString, Optional Prompt As String = vbNullString, Optional ShowLastUsed As Boolean = True)
  ShowRunDialogA hWnd, Title, Prompt, ShowLastUsed
End Sub

Public Sub AnimateWindow(Window As Object, Milliseconds As Long, Flags As ANIMATEWINDOW_Constants)
  AnimateWindowA Window, Milliseconds, Flags
End Sub

Public Sub ShowMouse()
  ShowMouseA
End Sub

Public Sub HideMouse()
  HideMouseA
End Sub

Public Function GetStartMode() As STARTMODE_Constants
  GetStartMode = GetStartModeA
End Function

Public Function SetDisplayMode(Width As Long, Height As Long, Bits As Long) As Boolean
  SetDisplayMode = SetDisplayModeA(Width, Height, Bits)
End Function

Public Function GetLastSystemShutdown() As DATE_AND_TIME
  GetLastSystemShutdown = GetLastSystemShutdownA
End Function

Public Function RefreshIconCache() As Boolean
  RefreshIconCache = RefreshIconCacheA
End Function
