VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileSystem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum SCANPATH_Constants
  ScanFloppyDrives = &H1
  ScanHardDrives = &H2
  ScanCDDrives = &H4
  ScanNetworkDrives = &H8
  ScanRamDrives = &H10
  ScanCustomPath = &H20
End Enum

Public Function FileMatchesWildCard(Path As String, WildCard As String) As Boolean
  FileMatchesWildCard = FileMatchesWildCardA(Path, WildCard)
End Function

Public Sub ScanPath(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional FindFiles As Boolean = True, Optional FindFolders As Boolean = True, Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False, Optional dWalk As DirWalk)
  ScanPathA RootPath, Path, Filter, FindFiles, FindFolders, SearchSubFolders, ForcePriority, dWalk
End Sub

Public Function GetSubFolders(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False) As Collection
  Set GetSubFolders = GetSubFoldersA(RootPath, Path, Filter, SearchSubFolders, ForcePriority)
End Function

Public Function GetSubFiles(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False) As Collection
  Set GetSubFiles = GetSubFilesA(RootPath, Path, Filter, SearchSubFolders, ForcePriority)
End Function

Public Function GetSubFoldersAndFiles(RootPath As SCANPATH_Constants, Optional Path As String, Optional Filter As String = "*.*", Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False) As Collection
  Set GetSubFoldersAndFiles = GetSubFoldersAndFilesA(RootPath, Path, Filter, SearchSubFolders, ForcePriority)
End Function

Public Function FolderContainsFiles(Path As String) As Boolean
  FolderContainsFiles = FolderContainsFilesA(Path)
End Function

Public Function FolderContainsFolders(Path As String) As Boolean
  FolderContainsFolders = FolderContainsFoldersA(Path)
End Function

Public Function FolderIsEmpty(Path As String) As Boolean
  FolderIsEmpty = FolderIsEmptyA(Path)
End Function
