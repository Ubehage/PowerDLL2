Attribute VB_Name = "modBrowser"
Option Explicit

Public Const BROWSER_KEY_DESKTOP = "/desktop/"
Public Const BROWSER_KEY_DOCUMENTS = "/mydocuments/"
Public Const BROWSER_KEY_MYCOMPUTER = "/mycomputer/"

Public Const BROWSER_NAME_MYCOMPUTER = "My Computer"
Public Const BROWSER_PATH_MYCOMPUTER = BROWSER_NAME_MYCOMPUTER

Public Function GetDesktopName() As String
  GetDesktopName = GetFileNameA(GetSpecialFolderPathA(sfDesktop))
End Function

Public Function GetDocumentsName() As String
  GetDocumentsName = GetFileNameA(GetSpecialFolderPathA(sfDocuments))
End Function

Public Function GetMyComputerName() As String
  GetMyComputerName = BROWSER_NAME_MYCOMPUTER
End Function

Public Function GetDesktopIcon() As IPictureDisp
  Set GetDesktopIcon = ExtractIconA(GetExplorerFile, 4, True)
End Function

Public Function GetMyDocumentsIcon() As IPictureDisp
  If IsWinXPA Then
    Set GetMyDocumentsIcon = ExtractIconA(GetShell32File, 127, True)
  Else
    Set GetMyDocumentsIcon = ExtractIconA(GetShell32File, 21, True)
  End If
End Function

Public Function GetMyComputerIcon() As IPictureDisp
  Set GetMyComputerIcon = ExtractIconA(GetExplorerFile, 1, True)
End Function

Public Function GetDriveCaption(Drive As PowerDLL2.Drive) As String
  Dim dCaption As String
  If Drive.DriveType = dtRemoveable Then
    dCaption = (GetDriveTypeName(Drive.DriveLetter) + " (")
  Else
    dCaption = (Drive.VolumeLabel + " (")
  End If
  dCaption = (dCaption + (UCase(Drive.DriveLetter) + ":)"))
  GetDriveCaption = dCaption
End Function

Public Function RenameFolder(Path As String, NewName As String) As Boolean
  Dim oName As String
  Dim nName As String
  oName = Path
  nName = NewName
  If GetFileNameA(nName) = "" Then
    nName = ((GetParentFolderA(oName) + "\") + nName)
  End If
  On Error GoTo Error1
  Name oName As nName
  RenameFolder = True
ExitNow:
  On Error GoTo 0
  Exit Function
Error1:
  Resume ExitNow
End Function

Public Function RenameFile(Path As String, NewName As String) As Boolean
  Dim oName As String
  Dim nName As String
  oName = Path
  nName = NewName
  If GetFileNameA(nName) = "" Then
    nName = ((GetParentFolderA(oName) + "\") + nName)
  End If
  On Error GoTo Error1
  Name oName As nName
  RenameFile = True
ExitNow:
  On Error GoTo 0
  Exit Function
Error1:
  Resume ExitNow
End Function

Public Function GetFoldersFromPath(Path As String) As Collection
  Dim fPath As String
  fPath = Path
  Set GetFoldersFromPath = New Collection
  Do Until fPath = ""
    GetFoldersFromPath.Add GetRootFolderFromPath(fPath)
    If GetFoldersFromPath.Item(GetFoldersFromPath.Count) = "" Then
      GetFoldersFromPath.Remove GetFoldersFromPath.Count
      GetFoldersFromPath.Add fPath
      fPath = ""
    Else
      fPath = Right(fPath, (Len(fPath) - (Len(GetFoldersFromPath.Item(GetFoldersFromPath.Count)) + 1)))
    End If
  Loop
End Function

Public Function GetRootFolderFromPath(Path As String) As String
  Dim i As Long
  For i = 1 To Len(Path)
    If Mid(Path, i, 1) = "\" Then
      GetRootFolderFromPath = Left(Path, (i - 1))
      Exit For
    End If
  Next
End Function

Public Function IsValidBrowserPath(Path As String) As Boolean
  If LCase(Path) = LCase(BROWSER_PATH_MYCOMPUTER) Then
    Path = BROWSER_PATH_MYCOMPUTER
    IsValidBrowserPath = True
  Else
    IsValidBrowserPath = IsFolderA(Path)
  End If
End Function

Public Function CanDeletePath(Path As String) As Boolean
  If Not LCase(Path) = LCase(BROWSER_PATH_MYCOMPUTER) Then
    If Not LCase(Path) = LCase(GetSpecialFolderPathA(sfDesktop)) Then
      If Not LCase(Path) = LCase(GetSpecialFolderPathA(sfDocuments)) Then
        CanDeletePath = True
      End If
    End If
  End If
End Function

Public Function GetPathFromBrowserString(CurrentPath As String, BrowserString As String) As String
  Dim bPath As String
  Dim bPaths As New Collection
  Dim nPath As String
  If Left(BrowserString, 2) = "\\" Then
    GetPathFromBrowserString = CurrentPath
    Exit Function
  End If
  bPath = BrowserString
  Do
    nPath = GetRootPath(bPath)
    If nPath = "" Then
      If bPath = "" Then
        Exit Do
      Else
        If bPath = "\" Then
          Exit Do
        End If
        nPath = bPath
        bPath = ""
      End If
    End If
    bPaths.Add nPath
    If Not bPath = "" Then
      bPath = Right(bPath, (Len(bPath) - (Len(nPath) + 1)))
    End If
  Loop
  If bPaths.Count = 0 Then
    GetPathFromBrowserString = CurrentPath
    Exit Function
  End If
  If Right(bPaths.Item(1), 1) = ":" Then
    nPath = bPaths.Item(1)
    bPaths.Remove 1
  Else
    nPath = CurrentPath
  End If
  Do Until bPaths.Count = 0
    If bPaths.Item(1) = ".." Then
      If nPath = BROWSER_PATH_MYCOMPUTER Then
        nPath = GetSpecialFolderPathA(sfDesktop)
      Else
        nPath = GetParentFolderA(nPath)
      End If
    Else
      If LCase(bPaths.Item(1)) = LCase(GetMyComputerName) Then
        nPath = BROWSER_PATH_MYCOMPUTER
      ElseIf LCase(bPaths.Item(1)) = LCase(GetDesktopName) Then
        nPath = GetSpecialFolderPathA(sfDesktop)
      ElseIf LCase(bPaths.Item(1)) = LCase(GetDocumentsName) Then
        nPath = GetSpecialFolderPathA(sfDocuments)
      Else
        If IsFolderA(((nPath + "\") + bPaths.Item(1))) Then
          nPath = ((nPath + "\") + bPaths.Item(1))
        End If
      End If
    End If
    bPaths.Remove 1
  Loop
  GetPathFromBrowserString = nPath
End Function

Private Function GetRootPath(Path As String) As String
  Dim i As Long
  For i = 1 To Len(Path)
    If Mid(Path, i, 1) = "\" Then
      GetRootPath = Left(Path, (i - 1))
      Exit For
    End If
  Next
End Function
