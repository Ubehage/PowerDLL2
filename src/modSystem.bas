Attribute VB_Name = "modSystem"
Option Explicit

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const KEY_CREATE_SUB_KEY  As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const READ_CONTROL As Long = &H20000
Private Const WRITE_DAC As Long = &H40000
Private Const WRITE_OWNER As Long = &H80000
Private Const SYNCHRONIZE As Long = &H100000
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Private Const KEY_READ As Long = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE As Long = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE As Long = KEY_READ

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_USERS As Long = &H80000003

Private Const REG_NONE As Long = &H0
Private Const REG_SZ As Long = &H1
Private Const REG_EXPAND_SZ As Long = &H2
Private Const REG_BINARY As Long = &H3
Private Const REG_DWORD As Long = &H4
Private Const REG_DWORD_LITTLE_ENDIAN As Long = &H4
Private Const REG_DWORD_BIG_ENDIAN As Long = &H5
Private Const REG_LINK As Long = &H6
Private Const REG_MULTI_SZ As Long = &H7
Private Const REG_RESOURCE_LIST As Long = &H8
Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = &H9
Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = &HA

Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_BADDB As Long = 1009
Private Const ERROR_BADKEY As Long = 1010
Private Const ERROR_CANTOPEN As Long = 1011
Private Const ERROR_CANTREAD As Long = 1012
Private Const ERROR_CANTWRITE As Long = 1013
Private Const ERROR_OUTOFMEMORY As Long = 14
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_MORE_DATA As Long = 234
Private Const ERROR_NO_MORE_ITEMS As Long = 259

Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_SETTINGCHANGE As Long = &H1A
Private Const SPI_SETNONCLIENTMETRICS As Long = &H2A
Private Const SMTO_ABORTIFHUNG As Long = &H2

Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SETWINDOWPOS = SWP_NOSIZE Or SWP_NOMOVE

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2

Private Const FORMAT_MESSSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSSAGE_ARGUMENT_ARRAY = &H2000

Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPI_UPDATEINIFILE = 1

Private Const PROP_PREVPROC = "PrevProc"
Private Const PROP_FORM = "FormObject"

Private Const SM_CLEANBOOT = 67

Private Type LUID
  dwLowPart As Long
  dwHighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
  udtLUID As LUID
  dwAttributes As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  laa As LUID_AND_ATTRIBUTES
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetDriveTypeB Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Long) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Public Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long

Public Declare Function IsPwrSuspendAllowed Lib "powrprof.dll" () As Long
Public Declare Function IsPwrHibernateAllowed Lib "powrprof.dll" () As Long

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ScreenToClientB Lib "user32" Alias "ScreenToClient" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreenB Lib "user32" Alias "ClientToScreen" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageID As Long, ByVal dwLanguageID As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function AnimateWindowAPI Lib "user32" Alias "AnimateWindow" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long

Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function apiOleTranslateColor Lib "oleaut32" Alias "OleTranslateColor" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long

Public Declare Function MulDiv Lib "kernel32" (ByVal Mul As Long, ByVal Nom As Long, ByVal Den As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal lpdwRes As Long, lpType As Long, lpData As Any, nSize As Long) As Long

Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)

Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Long
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Long) As Long

Public Function AnimateWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim lPrevProc As Long
  Dim lForm As Long
  Dim oForm As Form
  lPrevProc = GetProp(hWnd, PROP_PREVPROC)
  lForm = GetProp(hWnd, PROP_FORM)
  MoveMemory oForm, lForm, 4&
  Select Case Msg
    Case WM_PRINTCLIENT
      Dim tRect As RECT
      Dim hBrush As Long
      GetClientRect hWnd, tRect
      hBrush = CreateSolidBrush(OleTranslateColor(oForm.BackColor))
      FillRect wParam, tRect, hBrush
      DeleteObject hBrush
      If Not oForm.Picture Is Nothing Then
        Dim lSrcDC As Long
        Dim lMemDC As Long
        Dim lPrevBMP As Long
        lSrcDC = GetDC(0&)
        lMemDC = CreateCompatibleDC(lSrcDC)
        ReleaseDC 0, lSrcDC
        lPrevBMP = SelectObject(lMemDC, oForm.Picture.Handle)
        BitBlt wParam, 0, 0, HM2Pix(oForm.Picture.Width), HM2Pix(oForm.Picture.Height), lMemDC, 0, 0, vbSrcCopy
        SelectObject lMemDC, lPrevBMP
        DeleteDC lMemDC
      End If
  End Select
  MoveMemory oForm, 0&, 4&
  AnimateWindowProc = CallWindowProc(lPrevProc, hWnd, Msg, wParam, lParam)
End Function

Public Function OleTranslateColor(ByVal Color As Long) As Long
  apiOleTranslateColor Color, 0, OleTranslateColor
End Function

Public Function HM2Pix(ByVal Value As Long) As Long
  HM2Pix = (MulDiv(Value, 1440, 2540) / Screen.TwipsPerPixelX)
End Function

Public Sub WindowOnTopA(hWnd As Long, OnTop As Boolean)
  Dim wFlags As Long
  If OnTop Then
    wFlags = HWND_TOPMOST
  Else
    wFlags = HWND_NOTOPMOST
  End If
  SetWindowPos hWnd, wFlags, 0&, 0&, 0&, 0&, SWP_SETWINDOWPOS
End Sub

Public Function GetDriveTypeA(Drive As String) As DRIVETYPE_Constants
  Select Case GetDriveTypeB((Left(Drive, 1) + ":"))
    Case 2
      GetDriveTypeA = dtRemoveable
    Case 3
      GetDriveTypeA = dtHardDisk
    Case 4
      GetDriveTypeA = dtNetwork
    Case 5
      Select Case GetDriveTypeEx(Drive)
        Case FILE_DEVICE_CD_ROM
          GetDriveTypeA = dtCD
        Case FILE_DEVICE_DVD
          GetDriveTypeA = dtDVD
      End Select
    Case 6
      GetDriveTypeA = dtRamDrive
  End Select
End Function

Public Function GetDriveTypeEx(DriveLetter As String) As Long
  GetDriveTypeEx = GetDriveTypeB((Left(DriveLetter, 1) + ":"))
  If GetDriveTypeEx = dtCD Then
    If IsWinXPPlusA Then
      GetDriveTypeEx = GetMediaType(DriveLetter)
    Else
      GetDriveTypeEx = FILE_DEVICE_CD_ROM
    End If
  End If
End Function

Public Function GetMediaType(DriveLetter As String) As Long
  Dim sDrive As String
  Dim hDrive As Long
  Dim gmt As GET_MEDIA_TYPES
  Dim status As Long
  Dim returned As Long
  Dim mynull As Long
  If IsWinXPPlusA Then
    sDrive = (Left(DriveLetter, 1) + ":")
    hDrive = CreateFile("\\.\" & UCase(sDrive), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, mynull, OPEN_EXISTING, 0, mynull)
    If Not hDrive = INVALID_HANDLE_VALUE Then
      status = DeviceIoControl(hDrive, IOCTL_STORAGE_GET_MEDIA_TYPES_EX, mynull, 0, gmt, 2048, returned, ByVal 0)
      If Not status = 0 Then
        GetMediaType = gmt.DeviceType
      End If
    End If
    CloseHandle hDrive
  End If
End Function

Public Function IsProgramFileA(File As String) As Boolean
  Select Case LCase(GetFileExtensionA(File))
    Case "com", "exe", "bat"
      IsProgramFileA = True
  End Select
End Function

Public Function GetDriveLettersA() As String
  Dim dMask As Long
  Dim dMax As Long
  Dim dCount As Long
  Dim dLetters As String
  dMask = GetLogicalDrives
  If dMask Then
    dMax = Int((Log(dMask) / Log(2)))
    For dCount = 0 To dMax
      If ((2 ^ dMax) And dMask) Then
        If GetDriveTypeB((Chr((vbKeyA + dCount)) + ":")) > 1 Then
          dLetters = (dLetters + LCase(Chr((vbKeyA + dCount))))
        End If
      End If
    Next
  End If
  GetDriveLettersA = UCase(dLetters)
End Function

Public Function GetDriveObjectFromDriveLetterA(Drive As String) As Drive
  If DriveExistsA(Drive) Then
    Set GetDriveObjectFromDriveLetterA = New Drive
    GetDriveObjectFromDriveLetterA.DriveLetter = UCase(Left(Drive, 1))
  End If
End Function

Public Function EnableShutdownPrivileges() As Boolean
  Dim hProcessHandle As Long
  Dim hTokenHandle As Long
  Dim lpv_la As LUID
  Dim Token As TOKEN_PRIVILEGES
  hProcessHandle = GetCurrentProcess
  If Not hProcessHandle = 0 Then
    If Not OpenProcessToken(hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hTokenHandle) = 0 Then
      If Not LookupPrivilegeValue(vbNullString, "SeShutdownPrivilege", lpv_la) = 0 Then
        With Token
          .PrivilegeCount = 1
          With .laa
            .udtLUID = lpv_la
            .dwAttributes = SE_PRIVILEGE_ENABLED
          End With
        End With
        If Not AdjustTokenPrivileges(hTokenHandle, False, Token, ByVal 0&, ByVal 0&, ByVal 0&) = 0 Then
          EnableShutdownPrivileges = True
        End If
      End If
    End If
  End If
End Function

Public Function GetWindowsVersionA() As WINDOWS_VERSION_INFORMATION
  Dim osv As OSVERSIONINFO
  Dim pos As Integer
  Dim sVer As String
  Dim sBuild As String
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    GetWindowsVersionA.PlatformID = osv.PlatformID
    Select Case osv.PlatformID
      Case VER_PLATFORM_WIN32s
        GetWindowsVersionA.WindowsName = "Win32s"
      Case VER_PLATFORM_WIN32_NT
        GetWindowsVersionA.WindowsName = "Windows NT"
        Select Case osv.dwVerMajor
          Case 5
            Select Case osv.dwVerMinor
              Case 0
                GetWindowsVersionA.WindowsName = "Windows 2000"
              Case 1
                GetWindowsVersionA.WindowsName = "Windows XP"
            End Select
        End Select
      Case VER_PLATFORM_WIN32_WINDOWS
        Select Case osv.dwVerMinor
          Case 0
            GetWindowsVersionA.WindowsName = "Windows 95"
          Case 90
            GetWindowsVersionA.WindowsName = "Windows ME"
          Case Else
            GetWindowsVersionA.WindowsName = "Windows 98"
        End Select
    End Select
    GetWindowsVersionA.VersionNumber = ((Trim(Str(osv.dwVerMajor)) + ".") + Trim(Str(osv.dwVerMinor)))
    GetWindowsVersionA.BuildNumber = (osv.dwBuildNumber And &HFFFF&)
    pos = InStr(osv.szCSDVersion, vbNullChar)
    If pos Then
      GetWindowsVersionA.ServicePack = Left(osv.szCSDVersion, (pos - 1))
    End If
  End If
End Function

Public Function IsWin95A() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWin95A = ((osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (osv.dwVerMajor = 4 And osv.dwVerMinor = 0))
  End If
End Function

Public Function IsWin98A() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWin98A = ((osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And ((osv.dwVerMajor > 4 Or osv.dwVerMajor = 4) And osv.dwVerMinor > 0))
  End If
End Function

Public Function IsWinMEA() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWinMEA = ((osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (osv.dwVerMajor = 4 And osv.dwVerMinor = 90))
  End If
End Function

Public Function IsWinNT4A() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWinNT4A = ((osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor = 4))
  End If
End Function

Public Function IsWin2000A() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWin2000A = ((osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor = 5) And (osv.dwVerMinor = 0))
  End If
End Function

Public Function IsWinXPA() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWinXPA = ((osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor = 5) And (osv.dwVerMinor = 1))
  End If
End Function

Public Function IsWinXPPlusA() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWinXPPlusA = ((osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor >= 5 And osv.dwVerMinor >= 1))
  End If
End Function

Public Function IsWinNTA() As Boolean
  Dim osv As OSVERSIONINFO
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    IsWinNTA = ((osv.PlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwVerMajor >= 4))
  End If
End Function

Public Function GetWindowsNameA() As String
  GetWindowsNameA = GetWindowsVersionA.WindowsName
End Function

Public Function ExitWindowsA(ExitMode As EXITWINDOWS_Constants, Optional ForceFlag As Boolean = False) As Boolean
  Dim ewCode As Long
  Select Case ExitMode
    Case StandBy, Hibernate
      ewCode = 1
    Case Else
      ewCode = GetExitWindowsModeFromConstant(ExitMode)
  End Select
  If Not ewCode = EWX_NONE Then
    If PrepareForShutdown Then
      If (ExitMode = StandBy Or ExitMode = Hibernate) Then
        If ExitMode = StandBy Then
          ewCode = True
        Else
          ewCode = False
        End If
        SetSystemPowerState ewCode, ForceFlag
      Else
        If ForceFlag Then
          ewCode = (ewCode Or EWX_FORCE)
        End If
        ExitWindowsEx ewCode, 0&
      End If
      ExitWindowsA = True
    End If
  End If
End Function

Public Function CanStandByA() As Boolean
  CanStandByA = CBool(IsPwrSuspendAllowed)
End Function

Public Function CanHibernateA() As Boolean
  CanHibernateA = CBool(IsPwrHibernateAllowed)
End Function

Public Function GetMousePositionA() As POINTAPI
  GetCursorPos GetMousePositionA
End Function

Public Sub SetMousePositionA(NewPosition As POINTAPI)
  With NewPosition
    SetCursorPos .x, .y
  End With
End Sub

Public Sub ScreenToClientA(hWnd As Long, Position As POINTAPI)
  ScreenToClientB hWnd, Position
End Sub

Public Sub ClientToScreenA(hWnd As Long, Position As POINTAPI)
  ClientToScreenB hWnd, Position
End Sub

Public Sub StartCaptureA(hWnd As Long)
  SetCapture hWnd
End Sub

Public Sub EndCaptureA()
  ReleaseCapture
End Sub

Public Function GetErrorDescriptionA(ErrorCode As Long) As String
  Dim ret As Long
  Dim sBuff As String
  sBuff = Space(MAX_PATH)
  ret = FormatMessage(FORMAT_MESSSAGE_FROM_SYSTEM Or FORMAT_MESSSAGE_IGNORE_INSERTS Or FORMAT_MESSSAGE_MAX_WIDTH_MASK, 0&, ErrorCode, 0&, sBuff, Len(sBuff), 0&)
  If ret Then
    GetErrorDescriptionA = Left(sBuff, ret)
  End If
End Function

Public Sub SetDesktopWallpaperA(WallpaperFile As String)
  If Not WallpaperFile = "" Then
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0, WallpaperFile, SPI_UPDATEINIFILE
  End If
End Sub

Public Sub ShowRestartDialogA(Optional ExitMode As EXITWINDOWS_Constants = Restart, Optional hWnd As Long = 0&, Optional DialogMessage As String)
  If PrepareForShutdown Then
    MBRestartSystemDialog hWnd, DialogMessage, GetExitWindowsModeFromConstant(ExitMode)
  End If
End Sub

Public Sub ShowShutdownDialogA()
  If PrepareForShutdown Then
    MBShutdownDialog 0
  End If
End Sub

Public Sub ShowRunDialogA(Optional hWnd As Long = 0&, Optional Title As String = vbNullString, Optional Prompt As String = vbNullString, Optional ShowLastUsed As Boolean = True)
  If ShowLastUsed Then
    SHRunDialog hWnd, 0, 0, Title, Prompt, 0
  Else
    SHRunDialog hWnd, 0, 0, Title, Prompt, &H2
  End If
End Sub

Public Sub AnimateWindowA(Window As Object, Milliseconds As Long, Flags As ANIMATEWINDOW_Constants)
  SetProp Window.hWnd, PROP_PREVPROC, GetWindowLong(Window.hWnd, GWL_WNDPROC)
  SetProp Window.hWnd, PROP_FORM, ObjPtr(Window)
  SetWindowLong Window.hWnd, GWL_WNDPROC, AddressOf AnimateWindowProc
  AnimateWindowAPI Window.hWnd, Milliseconds, Flags
  SetWindowLong Window.hWnd, GWL_WNDPROC, GetProp(Window.hWnd, PROP_PREVPROC)
  RemoveProp Window.hWnd, PROP_PREVPROC
  RemoveProp Window.hWnd, PROP_FORM
  Window.Refresh
End Sub

Public Sub ShowMouseA()
  Do: Loop Until ShowCursor(1) > 0
End Sub

Public Sub HideMouseA()
  Do: Loop Until ShowCursor(0) < 0
End Sub

Public Function GetStartModeA() As STARTMODE_Constants
  GetStartModeA = GetSystemMetrics(SM_CLEANBOOT)
End Function

Public Function SetDisplayModeA(Width As Long, Height As Long, Bits As Long) As Boolean
  Dim dModes As clsDisplayModes
  Dim dMode As clsDisplayMode
  Set dModes = New clsDisplayModes
  Set dMode = dModes.GetDisplayMode(Width, Height, Bits)
  If Not dMode Is Nothing Then
    dMode.UseThisMode
    SetDisplayModeA = True
  End If
End Function

Public Function GetLastSystemShutdownA() As DATE_AND_TIME
  Dim hKey As Long
  Dim sKey As String
  Dim sValueName As String
  Dim ft As FILETIME
  Dim cbData As Long
  sKey = "System\CurrentControlSet\Control\Windows"
  sValueName = "ShutdownTime"
  If RegOpenKeyEx(RegObj.HKEY_LOCAL_MACHINE, sKey, 0&, KEY_READ, hKey) = ERROR_SUCCESS Then
    If Not hKey = 0 Then
      cbData = Len(ft)
      If RegQueryValueEx(hKey, sValueName, 0&, REG_BINARY, ft, cbData) = ERROR_SUCCESS Then
        GetLastSystemShutdownA = GetDateAndTimeFromFileTime(ft)
      End If
      RegCloseKey hKey
    End If
  End If
End Function

Public Function RefreshIconCacheA() As Boolean
  Dim hKey As Long
  Dim dwKeyType As Long
  Dim dwDataType As Long
  Dim dwDataSize As Long
  Dim sKeyName As String
  Dim sValue As String
  Dim sData As String
  Dim sDataRet As String
  Dim tmp As Long
  Dim sNewValue As String
  Dim dwNewValue As Long
  Dim Success As Long
  dwKeyType = HKEY_CURRENT_USER
  sKeyName = "Control Panel\Desktop\WindowMetrics"
  sValue = "Shell Icon Size"
  hKey = RegKeyOpen(HKEY_CURRENT_USER, sKeyName)
  If Not hKey = 0 Then
    'Debug.Print "RegKeyOpen: "; hKey
    dwDataSize = RegGetStringSize(ByVal hKey, sValue, dwDataType)
    'Debug.Print "RegGetStringSize: "; dwDataSize
    If dwDataSize > 0 Then
      sDataRet = RegGetStringValue(hKey, sValue, dwDataSize)
      If Not sDataRet = "" Then
        'Debug.Print "RegGetStringValue: "; sDataRet
        tmp = CLng(sDataRet)
        tmp = (tmp - 1)
        sNewValue = CStr(tmp) & Chr(0)
        dwNewValue = Len(sNewValue)
        If RegWriteStringValue(hKey, sValue, dwDataType, sNewValue) = ERROR_SUCCESS Then
          SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, Success
          sDataRet = sDataRet & Chr(0)
          RegWriteStringValue hKey, sValue, dwDataType, sDataRet
          SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, Success
        End If
      End If
    End If
  End If
  RegCloseKey hKey
End Function

Private Function RegGetStringSize(ByVal hKey As Long, ByVal sValue As String, dwDataType As Long) As Long
  Dim Success As Long
  Dim dwDataSize As Long
  Success = RegQueryValueEx(hKey, sValue, 0&, dwDataType, ByVal 0&, dwDataSize)
  If Success = ERROR_SUCCESS Then
    If dwDataType = REG_SZ Then
      RegGetStringSize = dwDataSize
    End If
  End If
End Function

Private Function RegKeyOpen(dwKeyType As Long, sKeyPath As String) As Long
  Dim hKey As Long
  Dim dwOptions As Long
  Dim sAttr As SECURITY_ATTRIBUTES
  sAttr.nLength = Len(sAttr)
  sAttr.bInheritHandle = False
  dwOptions = 0&
  If RegOpenKeyEx(dwKeyType, sKeyPath, dwOptions, KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
    RegKeyOpen = hKey
  End If
End Function

Private Function RegGetStringValue(ByVal hKey As Long, ByVal sValue As String, dwDataSize As Long) As String
  Dim sDataRet As String
  Dim dwDataRet As Long
  Dim Success As Long
  Dim pos As Long
  sDataRet = Space(dwDataSize)
  dwDataRet = Len(sDataRet)
  Success = RegQueryValueEx(hKey, sValue, ByVal 0&, dwDataSize, ByVal sDataRet, dwDataRet)
  If Success = ERROR_SUCCESS Then
    If dwDataRet > 0 Then
      pos = InStr(sDataRet, Chr(0))
      RegGetStringValue = Left(sDataRet, (pos - 1))
    End If
  End If
End Function

Private Function RegWriteStringValue(ByVal hKey, ByVal sValue, ByVal dwDataType, sNewValue) As Long
  Dim Success As Long
  Dim dwNewValue As Long
  dwNewValue = Len(sNewValue)
  If dwNewValue > 0 Then
    RegWriteStringValue = RegSetValueExString(hKey, sValue, 0&, dwDataType, sNewValue, dwNewValue)
  End If
End Function
