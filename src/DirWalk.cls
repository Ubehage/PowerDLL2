VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirWalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_FileCollection As Collection
Dim m_FolderCollection As Collection
Dim m_Collection As Collection

Public Event ItemFound(Path As String, Name As String, IsFile As Boolean, Cancel As Boolean)

Public Property Get FileCollection() As Collection
  Set FileCollection = m_FileCollection
End Property
Public Property Get FolderCollection() As Collection
  Set FolderCollection = m_FolderCollection
End Property
Public Property Get Collection() As Collection
  Set Collection = m_Collection
End Property
Public Property Let FileCollection(New_FileCollection As Collection)
  If Not m_FileCollection Is New_FileCollection Then
    Set m_FileCollection = New_FileCollection
  End If
End Property
Public Property Let FolderCollection(New_FolderCollection As Collection)
  If Not m_FolderCollection Is New_FolderCollection Then
    Set m_FolderCollection = New_FolderCollection
  End If
End Property
Public Property Let Collection(New_Collection As Collection)
  If Not m_Collection Is New_Collection Then
    Set m_Collection = New_Collection
  End If
End Property

Public Sub SearchPath(Path As String, Optional Filter As String = "*.*", Optional FindFiles As Boolean = True, Optional FindFolders As Boolean = True, Optional SearchSubFolders As Boolean = True, Optional ForcePriority As Boolean = False)
  Dim i As Long
  Dim sPath As String
  Dim DirsFound As Long
  Dim DirFound() As String
  Dim hItem As Long
  Dim fItem As WIN32_FIND_DATA
  Dim FoundItem As String
  sPath = FixPathA(Path)
  hItem = FindFirstFile((sPath + "\*.*"), fItem)
  CancelSearch = False
  If Not hItem = INVALID_HANDLE_VALUE Then
    Do
      FoundItem = TrimNull(fItem.cFileName)
      If (fItem.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
        If Not (FoundItem = "." Or FoundItem = "..") Then
          If (DirsFound Mod 10) = 0 Then
            ReDim Preserve DirFound(1 To (DirsFound + 10)) As String
          End If
          DirsFound = (DirsFound + 1)
          DirFound(DirsFound) = FoundItem
          If FindFolders Then
            If FileMatchesWildCardA(FoundItem, Filter) Then
              FoundNewItem sPath, FoundItem, False
            End If
          End If
        End If
      Else
        If FindFiles Then
          If FileMatchesWildCardA(FoundItem, Filter) Then
            FoundNewItem sPath, FoundItem, True
          End If
        End If
      End If
      If CancelSearch Then
        Exit Do
      End If
    Loop While FindNextFile(hItem, fItem)
  End If
  FindClose hItem
  If Not ForcePriority Then
    DoEvents
  End If
  If SearchSubFolders Then
    For i = 1 To DirsFound
      If CancelSearch Then
        Exit For
      End If
      SearchPath ((sPath + "\") + DirFound(i)), Filter, FindFiles, FindFolders, SearchSubFolders, ForcePriority
    Next
  End If
End Sub

Private Sub FoundNewItem(Path As String, Name As String, IsFile As Boolean)
  If Not m_Collection Is Nothing Then
    m_Collection.Add ((Path + "\") + Name)
  End If
  If IsFile Then
    If Not m_FileCollection Is Nothing Then
      m_FileCollection.Add ((Path + "\") + Name)
    End If
  Else
    If Not m_FolderCollection Is Nothing Then
      m_FolderCollection.Add ((Path + "\") + Name)
    End If
  End If
  RaiseEvent ItemFound(Path, Name, IsFile, CancelSearch)
End Sub

Private Sub Class_Terminate()
  Set m_FileCollection = Nothing
  Set m_FolderCollection = Nothing
  Set m_Collection = Nothing
End Sub
