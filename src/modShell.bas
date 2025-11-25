Attribute VB_Name = "basModShell"
Option Explicit

Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Long
    cbReserved2 As Long
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Public Type SHFILEOPSTRUCT
  hWnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAborted As Boolean
  hNameMaps As Long
  sProgress As String
End Type

Public Type CHOOSECOLORSTRUCT
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As Long
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Type OPENFILENAME
  nStructSize As Long
  hWndOwnder As Long
  hInstance As Long
  sFilter As String
  sCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  sFile As String
  nMaxFile As Long
  sFileTitle As String
  nMaxTitle As Long
  sInitialDir As String
  sDialogTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  sDefFileExt As String
  nCustData As Long
  fnHook As Long
  sTemplateName As String
  pvReserved As Long
  dwReserved As Long
  flagsEx As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&

Private Const INFINITE = -1&

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_VALIDATE = &H20
Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Private Const BIF_NONEWFOLDERBUTTON = &H200

Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_ANYCOLOR = &H100

Public Function GetAssociatedIconA(Path As String, Optional SmallIcon As Boolean = True) As IPictureDisp
  Dim hImg As Long
  Dim r As Long
  Dim fInfo As SHFILEINFO
  Dim shFlags As Long
  shFlags = BASIC_SHGFI_FLAGS
  If SmallIcon Then
    shFlags = shFlags Or SHGFI_SMALLICON
  Else
    shFlags = shFlags Or SHGFI_LARGEICON
  End If
  hImg = SHGetFileInfo(Path, 0&, fInfo, Len(fInfo), shFlags)
  Load frmObjects
  frmObjects.SizeIconPicture SmallIcon
  ImageList_Draw hImg, fInfo.iIcon, frmObjects.picIcon.hDC, 0, 0, ILD_TRANSPARENT
  frmObjects.picIcon.Picture = frmObjects.picIcon.Image
  Set GetAssociatedIconA = frmObjects.picIcon.Picture
  Unload frmObjects
  Set frmObjects = Nothing
End Function

Public Function GetIconsInFileA(File As String) As Long
  GetIconsInFileA = ExtractIconB(0&, File, -1)
End Function

Public Function ExtractIconA(File As String, IconIndex As Long, Optional SmallIcon As Boolean = True) As IPictureDisp
  Dim hIcon As Long
  Dim rIcon As IPictureDisp
  hIcon = ExtractIconB(0&, File, (IconIndex - 1))
  If Not hIcon = 0 Then
    Load frmObjects
    frmObjects.SizeIconPicture False
    DrawIcon frmObjects.picIcon.hDC, 0, 0, hIcon
    frmObjects.picIcon.Picture = frmObjects.picIcon.Image
    Set rIcon = frmObjects.picIcon.Picture
    Unload frmObjects
    Set frmObjects = Nothing
    DestroyIcon hIcon
    If SmallIcon Then
      Set ExtractIconA = ResizePictureA(rIcon, 16, 16)
    Else
      Set ExtractIconA = rIcon
    End If
    Set rIcon = Nothing
  End If
End Function

Public Function GetFileTypeDescriptionA(Path As String) As String
  Dim fInfo As SHFILEINFO
  SHGetFileInfo Path, 0&, fInfo, Len(fInfo), BASIC_SHGFI_FLAGS Or SHGFI_TYPENAME
  GetFileTypeDescriptionA = TrimNull(fInfo.szTypeName)
End Function

Public Function GetSpecialFolderPathA(SpecialPath As SPECIALFOLDER_Constants) As String
  Dim buff As String
  Dim Flags As Long
  buff = Space(MAX_LENGTH)
  If SHGetFolderPath(0&, SpecialPath Or Flags, -1, SHGFP_TYPE_CURRENT, buff) = S_OK Then
    GetSpecialFolderPathA = TrimNull(buff)
  End If
End Function

Public Function FileExistsA(Path As String) As Boolean
  Dim wfData As WIN32_FIND_DATA
  Dim hFile As Long
  If Not Path = "" Then
    hFile = FindFirstFile(Path, wfData)
    FileExistsA = Not (hFile = INVALID_HANDLE_VALUE)
    FindClose hFile
  End If
End Function

Public Function IsFolderA(Path As String)
  Dim wfData As WIN32_FIND_DATA
  Dim hFile As Long
  If Not Path = "" Then
    If Right(FixPathA(Path), 1) = ":" Then
      IsFolderA = DriveExistsA(Path)
    Else
      hFile = FindFirstFile(FixPathA(Path), wfData)
      If Not hFile = INVALID_HANDLE_VALUE Then
        If (wfData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
          IsFolderA = True
        End If
      End If
    End If
  End If
End Function

Public Function DriveExistsA(Drive As String) As Boolean
  DriveExistsA = Not (InStr(GetDriveLettersA, UCase(Left(Drive, 1))) = 0)
End Function

Public Function GetAssociatedFileA(Path As String) As String
  Dim aFile As String
  If IsProgramFileA(Path) Then
    GetAssociatedFileA = Path
  Else
    aFile = Space(MAX_PATH)
    Select Case FindExecutable(GetFileNameA(Path), (GetParentFolderA(Path) + "\"), aFile)
      Case Is >= 32
        aFile = Left(aFile, (InStr(aFile, vbNullChar) - 1))
      Case Else
        If IsFolderA(Path) Then
          aFile = GetExplorerFile
        ElseIf LCase(GetFileExtensionA(Path)) = "lnk" Then
          aFile = GetStartFile
        End If
    End Select
  End If
  GetAssociatedFileA = aFile
End Function

Public Function ExecuteFileA(Path As String, Optional Parameters As String, Optional WaitForExit As Boolean = False) As Boolean
  Dim eFile As String
  Dim eParam As String
  Dim cLine As String
  Dim pInfo As PROCESS_INFORMATION
  Dim sInfo As STARTUPINFO
  If Parameters = "" Then
    eFile = GetAssociatedFileA(Path)
    If Not eFile = Path Then
      ExecuteFileA = ExecuteFileA(eFile, QuotePathA(Path), WaitForExit)
      Exit Function
    End If
  End If
  eFile = Path
  If Not InStr(eFile, " ") = 0 Then
    eFile = QuotePathA(eFile)
  End If
  eParam = Parameters
  If (IsFolderA(eParam) Or FileExistsA(eParam)) Then
    If Not InStr(eParam, " ") = 0 Then
      eParam = QuotePathA(eParam)
    End If
  End If
  cLine = eFile
  If Not eParam = "" Then
    cLine = ((cLine + " ") + eParam)
  End If
  sInfo.cb = Len(sInfo)
  If Not CreateProcessA(0&, cLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, sInfo, pInfo) = 0 Then
    If WaitForExit Then
      WaitForSingleObject pInfo.hProcess, INFINITE
      CloseHandle pInfo.hProcess
      CloseHandle pInfo.hThread
    End If
    ExecuteFileA = True
  Else
    ExecuteFileA = False
  End If
End Function

Public Function ShellFileOperationA(hWnd As Long, shFiles As String, Optional shTarget As String, Optional shOperation As FILEOPERATION_Constants = foDelete, Optional shFlags As FILEOPERATION_FLAGS_Constants = FOF_ALLOWUNDO) As Boolean
  Dim sFiles As String
  Dim fStruct As SHFILEOPSTRUCT
  sFiles = Replace(shFiles, ";", vbNullChar)
  If Right(sFiles, 1) = vbNullChar Then
    If Not Mid(sFiles, (Len(sFiles) - 1), 1) = vbNullChar Then
      sFiles = (sFiles + vbNullChar)
    End If
  Else
    sFiles = (sFiles + String(2, vbNullChar))
  End If
  With fStruct
    .wFunc = shOperation
    .pFrom = sFiles
    .pTo = shTarget
    .fFlags = shFlags
    .hWnd = hWnd
  End With
  SHFileOperation fStruct
  ShellFileOperationA = Not fStruct.fAborted
End Function

Public Function BrowseForFolderA(Optional hWnd As Long = 0&, Optional Title As String = "Select Folder...", Optional RootFolder As SPECIALFOLDER_Constants = sfDesktop, Optional IncludeFiles As Boolean = False, Optional IncludeEditBox As Boolean = False, Optional IncludeNewFolderButton As Boolean = True) As String
  Dim bInfo As BROWSEINFO
  Dim pidl As Long
  Dim sPath As String
  If SHGetSpecialFolderLocation(hWnd, RootFolder, pidl) = S_OK Then
    With bInfo
      .hOwner = hWnd
      .pidlRoot = pidl
      .pszDisplayName = Space(MAX_PATH)
      .lpszTitle = Title
      .ulFlags = (BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE)
      If Not IncludeNewFolderButton Then
        .ulFlags = (.ulFlags Or BIF_NONEWFOLDERBUTTON)
      End If
      If IncludeFiles Then
        .ulFlags = (.ulFlags Or BIF_BROWSEINCLUDEFILES)
      End If
      If IncludeEditBox Then
        .ulFlags = (.ulFlags Or BIF_USENEWUI)
      End If
    End With
    pidl = SHBrowseForFolder(bInfo)
    If Not pidl = 0 Then
      sPath = Space(MAX_PATH)
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
        BrowseForFolderA = TrimNull(sPath)
      Else
        BrowseForFolderA = ("\\" + bInfo.pszDisplayName)
      End If
    End If
    CoTaskMemFree pidl
  End If
End Function

Public Function GetFileDateA(Path As String) As Date
  Dim fDate As String
  fDate = Trim(Str(FileDateTime(Path)))
  GetFileDateA = CDate(Left(fDate, (InStr(fDate, " ") - 1)))
End Function

Public Function GetFileTimeA(Path As String) As Date
  Dim fDate As String
  fDate = Trim(Str(FileDateTime(Path)))
  GetFileTimeA = CDate(Left(fDate, (InStr(fDate, " ") - 1)))
End Function

Public Sub AddAutoCompleteA(hWnd As Long, Flags As AUTOCOMPLETE_Flags, Optional UseAutoAppend As Boolean = True)
  Dim acFlags As Long
  acFlags = Flags
  If UseAutoAppend Then
    acFlags = (acFlags Or SHAC_AUTOAPPEND_ON)
  Else
    acFlags = (acFlags Or SHAC_AUTOAPPEND_OFF)
  End If
  SHAutoComplete hWnd, acFlags
End Sub

Public Function TrackPathFromStringA(PathString As String, CurrentPath As String) As String
  Dim nPath As String
  Dim pCol As New Collection
  If IsFullPath(PathString) Then
    TrackPathFromStringA = PathString
  Else
    nPath = CurrentPath
    pCol.Add PathString
    Do Until pCol.Item(pCol.Count) = ""
      pCol.Add GetParentFolderA(pCol.Item(pCol.Count))
    Loop
    pCol.Remove pCol.Count
    If Left(PathString, 1) = "\" Then
      nPath = Left(nPath, 2)
    End If
    Do Until pCol.Count = 0
      If GetFileNameA(pCol.Item(pCol.Count)) = ".." Then
        nPath = GetParentFolderA(nPath)
      ElseIf IsFolderInPathA(GetFileNameA(pCol.Item(pCol.Count)), nPath) Then
        nPath = ((nPath + "\") + GetFileNameA(pCol.Item(pCol.Count)))
      End If
      pCol.Remove pCol.Count
    Loop
    TrackPathFromStringA = nPath
  End If
End Function

Public Function IsFolderInPathA(Folder As String, Path As String) As Boolean
  Dim pCol As New Collection
  Dim fPath As String
  Dim NoFolder As Boolean
  fPath = Folder
  pCol.Add GetRootFolderFromPath(fPath)
  Do Until pCol.Item(pCol.Count) = ""
    fPath = Right(fPath, (Len(fPath) - (Len(pCol.Item(pCol.Count)) + 1)))
    pCol.Add GetRootFolderFromPath(fPath)
  Loop
  If Not (fPath = "" Or fPath = "\") Then
    pCol.Add fPath
  End If
  fPath = FixPathA(Path)
  Do Until pCol.Count = 0
    fPath = ((fPath + "\") + pCol.Item(1))
    If Not IsFolderA(fPath) Then
      NoFolder = True
      Exit Do
    End If
    pCol.Remove 1
  Loop
  IsFolderInPathA = Not NoFolder
End Function

Public Function SelectColorDialogA(hWnd As Long, Optional ShowFullDialog As Boolean = False, Optional HideCustomColors As Boolean = False, Optional SpecifyInitialColor As Boolean = False, Optional InitialColor As Long) As Long
  Dim cc As CHOOSECOLORSTRUCT
  With cc
    .Flags = CC_ANYCOLOR
    If ShowFullDialog Then
      .Flags = .Flags Or CC_FULLOPEN
    ElseIf HideCustomColors Then
      .Flags = .Flags Or CC_PREVENTFULLOPEN
    End If
    If SpecifyInitialColor Then
      .Flags = .Flags Or CC_RGBINIT
      .rgbResult = InitialColor
    End If
    .lStructSize = Len(cc)
    .hWndOwner = hWnd
  End With
  If ChooseColor(cc) = 1 Then
    SelectColorDialogA = cc.rgbResult
  End If
End Function

Public Function BrowseForFileA(hWnd As Long, Optional Title As String, Optional Filter As String, Optional FilterIndex As Long, Optional InitialPath As String, Optional InitialFile As String, Optional Save As Boolean, Optional Flags As OPENSAVE_FLAGS) As String
  Dim sFilter As String
  Dim sFile As String
  Dim sFlags As Long
  Dim oFN As OPENFILENAME
  sFlags = OFN_EXPLORER Or Flags
  With oFN
    .nStructSize = Len(oFN)
    .hWndOwnder = hWnd
    sFilter = Replace(Filter, "|", vbNullChar)
    If Not Right(sFilter, 2) = (vbNullChar + vbNullChar) Then
      If Right(sFilter, 1) = vbNullChar Then
        sFilter = (sFilter + vbNullChar)
      Else
        sFilter = (sFilter + (vbNullChar + vbNullChar))
      End If
    End If
    .sFilter = sFilter
    .nFilterIndex = FilterIndex
    If IsMissing(InitialFile) Then
      If Save Then
        sFile = (Space(1024) + (vbNullChar + vbNullChar))
      Else
        sFile = (Space(4096) + (vbNullChar + vbNullChar))
      End If
    Else
      If Save Then
        sFile = (GetFileNameA(InitialFile) + (Space(1024) + (vbNullChar + vbNullChar)))
      Else
        sFile = (GetFileNameA(InitialFile) + (Space(4096) + (vbNullChar + vbNullChar)))
      End If
    End If
    .sFile = sFile
    .nMaxFile = Len(.sFile)
    .sFileTitle = ((vbNullChar + Space(512)) + (vbNullChar + vbNullChar))
    .nMaxTitle = Len(.sFileTitle)
    If Not IsMissing(InitialPath) Then
      .sInitialDir = (InitialPath + (vbNullChar + vbNullChar))
    End If
    If Not IsMissing(Title) Then
      .sDialogTitle = Title
    End If
    .Flags = sFlags
  End With
  If Save Then
    If GetSaveFileName(oFN) Then
      sFile = oFN.sFile
    End If
  Else
    If GetOpenFileName(oFN) Then
      sFile = Trim(oFN.sFile)
    End If
  End If
  While Right(sFile, 1) = vbNullChar
    sFile = Left(sFile, (Len(sFile) - 1))
  Wend
  While Right(sFile, 1) = " "
    sFile = Left(sFile, (Len(sFile) - 1))
  Wend
  While Right(sFile, 1) = vbNullChar
    sFile = Left(sFile, (Len(sFile) - 1))
  Wend
  sFile = Replace(sFile, vbNullChar, ";")
  BrowseForFileA = sFile
End Function
