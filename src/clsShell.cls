VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShell"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum SPECIALFOLDER_Constants
  sfDesktop = CSIDL_DESKTOP
  sfInternetExplorer = CSIDL_INTERNET
  sfPrograms = CSIDL_PROGRAMS
  sfControlPanel = CSIDL_CONTROLS
  sfPrinters = CSIDL_PRINTERS
  sfDocuments = CSIDL_PERSONAL
  sfFavorites = CSIDL_FAVORITES
  sfStartup = CSIDL_STARTUP
  sfRecent = CSIDL_RECENT
  sfSendTo = CSIDL_SENDTO
  sfRecycle = CSIDL_BITBUCKET
  sfStartMenu = CSIDL_STARTMENU
  sfMyDocuments = CSIDL_MYDOCUMENTS
  sfMyMusic = CSIDL_MYMUSIC
  sfMyVideo = CSIDL_MYVIDEO
  sfDesktopDirectory = CSIDL_DESKTOPDIRECTORY
  sfMyComputer = CSIDL_DRIVES
  sfNetwork = CSIDL_NETWORK
  sfNethood = CSIDL_NETHOOD
  sfFonts = CSIDL_FONTS
  sfTemplates = CSIDL_TEMPLATES
  sfCommonStartMenu = CSIDL_COMMON_STARTMENU
  sfCommonPrograms = CSIDL_COMMON_PROGRAMS
  sfCommonStartup = CSIDL_COMMON_STARTUP
  sfCommonDesktopDirectory = CSIDL_COMMON_DESKTOP_DIRECTORY
  sfApplicationData = CSIDL_APPDATA
  sfPrintHood = CSIDL_PRINTHOOD
  sfAlternateStartup = CSIDL_ALTSTARTUP
  sfCommonAlternateStartup = CSIDL_COMMON_ALTSTARTUP
  sfCommonFavorites = CSIDL_COMMON_FAVORITES
  sfTemporaryInternetFiles = CSIDL_INTERNET_CACHE
  sfCookies = CSIDL_COOKIES
  sfInternetHistory = CSIDL_HISTORY
  sfCommonApplicationData = CSIDL_COMMON_APPDATA
  sfWindowsDirectory = CSIDL_WINDOWS
  sfSystemDirectory = CSIDL_SYSTEM
  sfProgramFiles = CSIDL_PROGRAM_FILES
  sfMyPictures = CSIDL_MYPICTURES
  sfProfile = CSIDL_PROFILE
  sfProgramFilesCommon = CSIDL_PROGRAM_FILES_COMMON
  sfCommonTemplates = CSIDL_COMMON_TEMPLATES
  sfCommonDocuments = CSIDL_COMMON_DOCUMENTS
  sfCommonAdminTools = CSIDL_COMMON_ADMINTOOLS
  sfAdminTools = CSIDL_ADMINTOOLS
  sfConnections = CSIDL_CONNECTIONS
  sfCommonMusic = CSIDL_COMMON_MUSIC
  sfCommonPictures = CSIDL_COMMON_PICTURES
  sfCommonVideo = CSIDL_COMMON_VIDEO
  sfResources = CSIDL_RESOURCES
  sfLocalizedResources = CSIDL_RESOURCES_LOCALIZED
  sfOEMLinks = CSIDL_COMMON_OEM_LINKS
  sfCDBurn = CSIDL_CDBURN_AREA
  sfComputersNearMe = CSIDL_COMPUTERSNEARME
End Enum

Public Enum FILEOPERATION_Constants
  foCopy = FO_COPY
  foDelete = FO_DELETE
  foMove = FO_MOVE
  foRename = FO_RENAME
End Enum

Public Enum FILEOPERATION_FLAGS_Constants
  fofMultipleDestFiles = FOF_MULTIDESTFILES
  fofConfirmMouse = FOF_CONFIRMMOUSE
  fofSilent = FOF_SILENT
  fofRenameOnCollision = FOF_RENAMEONCOLLISION
  fofNoConfirmation = FOF_NOCONFIRMATION
  fofWantMappingHandle = FOF_WANTMAPPINGHANDLE
  fofAllowUndo = FOF_ALLOWUNDO
  fofFilesOnly = FOF_FILESONLY
  fofSimpleProgress = FOF_SIMPLEPROGRESS
  fofNoConfirmMkDir = FOF_NOCONFIRMMKDIR
  fofNoErrorUI = FOF_NOERRORUI
  fofNoCopySecurityAttribs = FOF_NOCOPYSECURITYATTRIBS
  fofNoRecursion = FOF_NORECURSION
  fofNoConnectedElements = FOF_NO_CONNECTED_ELEMENTS
  fofWantNukeWarning = FOF_WANTNUKEWARNING
  fofNoRecurseReparse = FOF_NORECURSEREPARSE
End Enum

Public Enum AUTOCOMPLETE_Flags
  Default = SHAC_DEFAULT
  FileSystem = SHAC_FILESYSTEM
  URLHistory = SHAC_URLHISTORY
  URLMRU = SHAC_URLMRU
  URLAll = SHAC_URLALL
  UseTabCompletion = SHAC_USETAB
End Enum

Public Enum OPENSAVE_FLAGS
  osAllowMultiSelect = OFN_ALLOWMULTISELECT
  osCreatePromptForNonexistingFile = OFN_CREATEPROMPT
  osFileMustExist = OFN_FILEMUSTEXIST
  osHideReadOnly = OFN_HIDEREADONLY
  osNoChangeDir = OFN_NOCHANGEDIR
  osPromptOverwrite = OFN_OVERWRITEPROMPT
  osPathMustExist = OFN_PATHMUSTEXIST
End Enum

Public Function GetAssociatedIcon(Path As String, Optional SmallIcon As Boolean = True) As IPictureDisp
  Set GetAssociatedIcon = GetAssociatedIconA(Path, SmallIcon)
End Function

Public Function GetIconsInFile(File As String) As Long
  GetIconsInFile = GetIconsInFileA(File)
End Function

Public Function ExtractIcon(File As String, IconIndex As Long, Optional SmallIcon As Boolean = True) As IPictureDisp
  Set ExtractIcon = ExtractIconA(File, IconIndex, SmallIcon)
End Function

Public Function GetFileTypeDescription(Path As String) As String
  GetFileTypeDescription = GetFileTypeDescriptionA(Path)
End Function

Public Function GetSpecialFolderPath(SpecialPath As SPECIALFOLDER_Constants) As String
  GetSpecialFolderPath = GetSpecialFolderPathA(SpecialPath)
End Function

Public Function FileExists(Path As String) As Boolean
  FileExists = FileExistsA(Path)
End Function

Public Function IsFolder(Path As String) As Boolean
  IsFolder = IsFolderA(Path)
End Function

Public Function DriveExists(Drive As String) As Boolean
  DriveExists = DriveExistsA(Drive)
End Function

Public Function GetAssociatedFile(Path As String) As String
  GetAssociatedFile = GetAssociatedFileA(Path)
End Function

Public Function ExecuteFile(Path As String, Optional Parameters As String, Optional WaitForExit As Boolean = False) As Boolean
  ExecuteFile = ExecuteFileA(Path, Parameters, WaitForExit)
End Function

Public Function ShellFileOperation(hWnd As Long, shFiles As String, Optional shTarget As String, Optional shOperation As FILEOPERATION_Constants = foDelete, Optional shFlags As FILEOPERATION_FLAGS_Constants = FOF_ALLOWUNDO) As Boolean
  ShellFileOperation = ShellFileOperationA(hWnd, shFiles, shTarget, shOperation, shFlags)
End Function

Public Function BrowseForFolder(Optional hWnd As Long = 0&, Optional Title As String = "Select Folder...", Optional RootFolder As SPECIALFOLDER_Constants = sfDesktop, Optional IncludeFiles As Boolean = False, Optional IncludeEditBox As Boolean = False, Optional IncludeNewFolderButton As Boolean = True) As String
  BrowseForFolder = BrowseForFolderA(hWnd, Title, RootFolder, IncludeFiles, IncludeEditBox, IncludeNewFolderButton)
End Function

Public Function GetFileDate(Path As String) As Date
  GetFileDate = GetFileDateA(Path)
End Function

Public Function GetFileTime(Path As String) As Date
  GetFileTime = GetFileTimeA(Path)
End Function

Public Sub AddAutoComplete(hWnd As Long, Flags As AUTOCOMPLETE_Flags, Optional UseAutoAppend As Boolean = True)
  AddAutoCompleteA hWnd, Flags, UseAutoAppend
End Sub

Public Function TrackPathFromString(PathString As String, CurrentPath As String) As String
  TrackPathFromString = TrackPathFromStringA(PathString, CurrentPath)
End Function

Public Function IsFolderInPath(Folder As String, Path As String) As Boolean
  IsFolderInPath = IsFolderInPathA(Folder, Path)
End Function

Public Function SelectColorDialog(hWnd As Long, Optional ShowFullDialog As Boolean = False, Optional HideCustomColors As Boolean = False, Optional SpecifyInitialColor As Boolean = False, Optional InitialColor As Long) As Long
  SelectColorDialog = SelectColorDialogA(hWnd, ShowFullDialog, HideCustomColors, SpecifyInitialColor, InitialColor)
End Function

Public Function BrowseForFile(hWnd As Long, Optional Title As String, Optional Filter As String, Optional FilterIndex As Long, Optional InitialPath As String, Optional InitialFile As String, Optional Save As Boolean, Optional Flags As OPENSAVE_FLAGS) As String
  BrowseForFile = BrowseForFileA(hWnd, Title, Filter, FilterIndex, InitialPath, InitialFile, Save, Flags)
End Function
