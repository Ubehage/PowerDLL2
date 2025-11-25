Attribute VB_Name = "modShellChange1"
Option Explicit

Private Const MAX_PATH_LEN = 260

Public Type SHFILEINFOBYTE
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName(1 To MAX_PATH_LEN) As Byte
  szTypeName(1 To 80) As Byte
End Type

Public Type PIDLSTRUCT
  pidl As Long
  bWatchSubFolders As Long
End Type

Public Type SHNOTIFYSTRUCT
  dwItem1 As Long
  dwItem2 As Long
End Type

Global m_PIDLDesktop As Long

Public Function GetPIDLFromFolderID(hOwner As Long, nFolder As Long) As Long
  Dim pidl As Long
  If SHGetSpecialFolderLocation(hOwner, nFolder, pidl) = ERROR_SUCCESS Then
    GetPIDLFromFolderID = pidl
  End If
End Function

Public Function GetDisplayNameFromPIDL(pidl As Long) As String
  Dim sfib As SHFILEINFOBYTE
  If SHGetFileInfoPidl(pidl, 0, sfib, Len(sfib), SHGFI_PIDL Or SHGFI_DISPLAYNAME) Then
    GetDisplayNameFromPIDL = GetStrFromBufferA(StrConv(sfib.szDisplayName, vbUnicode))
  End If
End Function

Public Function GetPathFromPIDL(pidl As Long) As String
  Dim sPath As String * MAX_PATH
  If SHGetPathFromIDList(pidl, sPath) Then
    GetPathFromPIDL = GetStrFromBufferA(sPath)
  End If
End Function

Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left(sz, (InStr(sz, vbNullChar) - 1))
  Else
    GetStrFromBufferA = sz
  End If
End Function
