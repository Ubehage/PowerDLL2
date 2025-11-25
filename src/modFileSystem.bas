Attribute VB_Name = "modFileSystem"
Option Explicit

Global CancelSearch As Boolean

Public Function FileMatchesWildCardA(Path As String, WildCard As String) As Boolean
  Dim i As Long
  Dim fName As String
  Dim fCard As String
  fName = GetFileNameA(Path)
  If fName = "" Then
    fName = Path
  End If
  For i = 1 To Len(WildCard)
    If Mid(WildCard, i, 1) = ";" Then
      If FileMatchesWildCardA(fName, fCard) Then
        FileMatchesWildCardA = True
        Exit Function
      End If
      fCard = ""
    Else
      fCard = (fCard + Mid(WildCard, i, 1))
    End If
  Next
  If fCard = "" Then
    fCard = "*"
  End If
  If InStr(fName, ".") = 0 Then
    fName = (fName + ".*")
  ElseIf Right(fName, 1) = "." Then
    fName = (fName + "*")
  End If
  If InStr(fCard, ".") = 0 Then
    fCard = (fCard + ".*")
  ElseIf Right(fCard, 1) = "." Then
    fCard = (fCard + "*")
  End If
  FileMatchesWildCardA = LCase(fName) Like LCase(fCard)
End Function

Public Sub ScanPathA(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional FindFiles As Boolean = True, Optional FindFolders As Boolean = True, Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False, Optional dWalk As DirWalk)
  Dim sFloppy As Boolean
  Dim sHardDisk As Boolean
  Dim sCD As Boolean
  Dim sNetwork As Boolean
  Dim sRam As Boolean
  Dim sCustom As Boolean
  Dim sThis As Boolean
  Dim dW As DirWalk
  Dim i As Long
  Dim d As Drives
  sFloppy = (RootPath And ScanFloppyDrives)
  sHardDisk = (RootPath And ScanHardDrives)
  sCD = (RootPath And ScanCDDrives)
  sNetwork = (RootPath And ScanNetworkDrives)
  sRam = (RootPath And ScanRamDrives)
  sCustom = (RootPath And ScanCustomPath)
  If dWalk Is Nothing Then
    Set dW = New DirWalk
  Else
    Set dW = dWalk
  End If
  CancelSearch = False
  If sCustom Then
    If IsFolderA(Path) Then
      dW.SearchPath Path, Filter, FindFiles, FindFolders, SearchSubFolders, ForcePriority
    End If
  End If
  Set d = New Drives
  For i = 1 To d.Count
    If CancelSearch Then
      Exit For
    End If
    sThis = False
    With d.Drive(i)
      If sFloppy Then
        If .DriveType = dtRemoveable Then
          sThis = True
        End If
      End If
      If sHardDisk Then
        If .DriveType = dtHardDisk Then
          sThis = True
        End If
      End If
      If sCD Then
        If .DriveType = dtCD Then
          sThis = True
        End If
      End If
      If sNetwork Then
        If .DriveType = dtNetwork Then
          sThis = True
        End If
      End If
      If sRam Then
        If .DriveType = dtRamDrive Then
          sThis = True
        End If
      End If
      If sThis Then
        dW.SearchPath (.DriveLetter + ":"), Filter, FindFiles, FindFolders, SearchSubFolders, ForcePriority
      End If
    End With
  Next
End Sub

Public Function GetSubFoldersA(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False) As Collection
  Dim dWalk As DirWalk
  Set dWalk = New DirWalk
  Set GetSubFoldersA = New Collection
  dWalk.FolderCollection = GetSubFoldersA
  ScanPathA RootPath, Path, Filter, False, True, SearchSubFolders, ForcePriority, dWalk
End Function

Public Function GetSubFilesA(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False) As Collection
  Dim dWalk As DirWalk
  Set dWalk = New DirWalk
  Set GetSubFilesA = New Collection
  dWalk.FileCollection = GetSubFilesA
  ScanPathA RootPath, Path, Filter, True, False, SearchSubFolders, ForcePriority, dWalk
End Function

Public Function GetSubFoldersAndFilesA(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False) As Collection
  Dim dWalk As DirWalk
  Set dWalk = New DirWalk
  Set GetSubFoldersAndFilesA = New Collection
  dWalk.Collection = GetSubFoldersAndFilesA
  ScanPathA RootPath, Path, Filter, True, True, SearchSubFolders, ForcePriority, dWalk
End Function

Public Function FolderContainsFilesA(Path As String) As Boolean
  Dim wfData As WIN32_FIND_DATA
  Dim hFile As Long
  hFile = FindFirstFile((FixPathA(Path) + "\*.*"), wfData)
  If Not hFile = INVALID_HANDLE_VALUE Then
    Do
      If (Not (wfData.dwFileAttributes And vbDirectory) = vbDirectory) Then
        FolderContainsFilesA = True
        Exit Do
      End If
    Loop
  End If
  FindClose hFile
End Function

Public Function FolderContainsFoldersA(Path As String) As Boolean
  Dim wfData As WIN32_FIND_DATA
  Dim hFile As Long
  hFile = FindFirstFile((FixPathA(Path) + "\*.*"), wfData)
  If Not hFile = INVALID_HANDLE_VALUE Then
    Do
      If (wfData.dwFileAttributes And vbDirectory) Then
        If Not (Left(wfData.cFileName, 1) = "." Or Left(wfData.cFileName, 2) = "..") Then
          FolderContainsFoldersA = True
          Exit Do
        End If
      End If
    Loop While FindNextFile(hFile, wfData)
  End If
  FindClose hFile
End Function

Public Function FolderIsEmptyA(Path As String) As Boolean
  If Not FolderContainsFilesA(Path) Then
    If Not FolderContainsFoldersA(Path) Then
      FolderIsEmptyA = True
    End If
  End If
End Function
