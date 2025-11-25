VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const FOLDER_PREFIX = "fo"
Private Const FILE_PREFIX = "fi"

Dim WithEvents m_ListViewParent As MSComctlLib.ListView
Attribute m_ListViewParent.VB_VarHelpID = -1
Dim m_Path As String
Dim m_ShowFolders As Boolean
Dim m_ShowFiles As Boolean
Dim m_AllowRename As Boolean
Dim m_DirShow As PowerDLL2.DirShow

Dim imlIcons As MSComctlLib.ImageList

Dim DiskRename As Boolean

Dim SelItem As MSComctlLib.ListItem

Public Event FolderClick(Path As String)
Public Event FolderRenamed(OldPath As String, NewPath As String)
Public Event FolderRightClick(Path As String)
Public Event FileClick(Path As String)
Public Event FileRenamed(OldPath As String, NewPath As String)
Public Event FileRightClick(Path As String)
Public Event FileDblClick(Path As String)
Public Event RightClick()
Public Event PathChange(Path As String)
Public Event KeyDown(KeyCode As VBRUN.KeyCodeConstants, Shift As VBRUN.ShiftConstants)
Public Event ExecutingFile(Path As String, Cancel As Boolean)

Public Property Get ListViewParent() As Object
  Set ListViewParent = m_ListViewParent
End Property
Public Property Get Path() As String
  Path = m_Path
End Property
Public Property Get ShowFolders() As Boolean
  ShowFolders = m_ShowFolders
End Property
Public Property Get ShowFiles() As Boolean
  ShowFiles = m_ShowFiles
End Property
Public Property Get AllowRename() As Boolean
  AllowRename = m_AllowRename
End Property
Public Property Get MultipleSelected() As Boolean
  MultipleSelected = IsMultipleSelected
End Property
Public Property Get SelectedItems() As String
  SelectedItems = GetSelectedItems
End Property
Public Property Get SelectedItem() As String
  If CanRun Then
    If Not m_ListViewParent.SelectedItem Is Nothing Then
      SelectedItem = GetItemPath(m_ListViewParent.SelectedItem)
    End If
  End If
End Property
Public Property Get DirShow() As PowerDLL2.DirShow
  Set DirShow = m_DirShow
End Property
Public Property Let ListViewParent(New_ListViewParent As Object)
  If Not m_ListViewParent Is New_ListViewParent Then
    Set m_ListViewParent = New_ListViewParent
    PrepareListView
    Refresh
  End If
End Property
Public Property Let Path(New_Path As String)
  If Not m_Path = New_Path Then
    If IsValidPath(New_Path) Then
      m_Path = New_Path
      Refresh
      UpdateDirShow
    End If
  End If
End Property
Public Property Let ShowFolders(New_ShowFolders As Boolean)
  If Not m_ShowFolders = New_ShowFolders Then
    m_ShowFolders = New_ShowFolders
    Refresh
  End If
End Property
Public Property Let ShowFiles(New_ShowFiles As Boolean)
  If Not m_ShowFiles = New_ShowFiles Then
    m_ShowFiles = New_ShowFiles
    Refresh
  End If
End Property
Public Property Let AllowRename(New_AllowRename As Boolean)
  If Not m_AllowRename = New_AllowRename Then
    m_AllowRename = New_AllowRename
    SetAllowRename
  End If
End Property
Public Property Let DirShow(New_DirShow As PowerDLL2.DirShow)
  If Not m_DirShow Is New_DirShow Then
    Set m_DirShow = New_DirShow
    If Not m_DirShow Is Nothing Then
      m_DirShow.FileShow = Me
    End If
    UpdateDirShow
  End If
End Property

Public Sub Refresh()
  FillListView
  UpdateDirShow
End Sub

Public Sub Detach()
  ClearListView
  ClearColumnHeaders
End Sub

Public Sub GoToPath(NewPath As String)
  m_Path = GetPathFromBrowserString(m_Path, NewPath)
  Refresh
  'If IsValidBrowserPath(NewPath) Then
    'Path = NewPath
  'Else
    'Select Case LCase(NewPath)
      'Case LCase(GetDesktopName)
        'Path = GetSpecialFolderPathA(sfDesktop)
      'Case LCase(GetMyComputerName)
        'Path = BROWSER_PATH_MYCOMPUTER
      'Case Else
        'If IsFolderInPathA(NewPath, m_Path) Then
          'Path = ((FixPathA(m_Path) + "\") + NewPath)
        'End If
    'End Select
  'End If
End Sub

Private Sub ClearColumnHeaders()
  If CanRun Then
    m_ListViewParent.ColumnHeaders.Clear
  End If
End Sub

Private Sub PrepareListView()
  If CanRun Then
    With m_ListViewParent
      .View = lvwReport
      .Arrange = lvwAutoLeft
      .MultiSelect = True
      .HideSelection = False
      .HideColumnHeaders = False
      With .ColumnHeaders
        .Add , "Name"
        .Add , "Path"
        .Add , "Size"
        .Add , "Type"
        .Add , "Date"
      End With
      .OLEDragMode = ccOLEDragAutomatic
      .OLEDropMode = ccOLEDropManual
    End With
    SetAllowRename
  End If
End Sub

Private Sub SetAllowRename()
  If CanRun Then
    If m_AllowRename Then
      m_ListViewParent.LabelEdit = lvwAutomatic
    Else
      m_ListViewParent.LabelEdit = lvwManual
    End If
  End If
End Sub

Private Sub ClearListView()
  If CanRun Then
    m_ListViewParent.ListItems.Clear
    ClearImageList
  End If
End Sub

Private Sub ClearImageList()
  If CanRun Then
    If m_ListViewParent.SmallIcons Is Nothing Then
      Set imlIcons = Nothing
    Else
      Set imlIcons = m_ListViewParent.SmallIcons
      Set m_ListViewParent.SmallIcons = Nothing
      With imlIcons
        .ListImages.Clear
        .ImageWidth = 16
        .ImageHeight = 16
        .ListImages.Add , , LoadResPicture(101, vbResIcon)
      End With
      Set m_ListViewParent.SmallIcons = imlIcons
    End If
  End If
End Sub

Private Function UsingIcons() As Boolean
  UsingIcons = Not (imlIcons Is Nothing)
End Function

Private Function CanRun() As Boolean
  CanRun = Not (m_ListViewParent Is Nothing)
End Function

Private Sub SetColumnHeaderNames(Optional UsingMyComputer As Boolean = False)
  If UsingMyComputer Then
    With m_ListViewParent.ColumnHeaders
      .Item(1).Text = "Name"
      .Item(2).Text = "Type"
      .Item(3).Text = "Total Size"
      .Item(4).Text = "Free Space"
      .Item(5).Text = "Comments"
    End With
  Else
    With m_ListViewParent.ColumnHeaders
      .Item(1).Text = "Name"
      .Item(2).Text = "Path"
      .Item(3).Text = "Size"
      .Item(4).Text = "Type"
      .Item(5).Text = "Date"
    End With
  End If
End Sub

Private Sub FillListView()
  ClearListView
  If m_Path = BROWSER_PATH_MYCOMPUTER Then
    AddMyComputerItems
  Else
    If LCase(m_Path) = LCase(GetSpecialFolderPathA(sfDesktop)) Then
      AddDesktopItems
    Else
      AddPathItems
    End If
  End If
End Sub

Private Sub AddPathItems()
  SetColumnHeaderNames False
  If m_ShowFolders Then
    With SortCollectionA(GetSubFoldersA(ScanCustomPath, m_Path, , False, True))
      Do Until .Count = 0
        AddFolderItem .Item(1)
        .Remove 1
      Loop
    End With
  End If
  If m_ShowFiles Then
    With SortCollectionA(GetSubFilesA(ScanCustomPath, m_Path, , False, True))
      Do Until .Count = 0
        AddFileItem .Item(1)
        .Remove 1
      Loop
    End With
  End If
End Sub

Private Sub AddDesktopItems()
  SetColumnHeaderNames False
  If m_ShowFolders Then
    With m_ListViewParent.ListItems.Add(, (FOLDER_PREFIX + GetSpecialFolderPathA(sfDocuments)), GetDocumentsName)
      .SubItems(1) = GetDesktopName
      .SubItems(2) = GetByteStringA(0)
      .SubItems(3) = GetFileTypeDescriptionA(GetSpecialFolderPathA(sfDocuments))
      .SubItems(4) = Trim(Str(GetFileDateA(GetSpecialFolderPathA(sfDocuments))))
      If UsingIcons Then
        .SmallIcon = GetIconIndex(GetMyDocumentsIcon)
      End If
    End With
    With m_ListViewParent.ListItems.Add(, BROWSER_KEY_MYCOMPUTER, GetMyComputerName)
      .SubItems(1) = GetDesktopName
      .SubItems(2) = GetByteStringA(0)
      .SubItems(3) = ""
      .SubItems(4) = ""
      If UsingIcons Then
        .SmallIcon = GetIconIndex(GetMyComputerIcon)
      End If
    End With
    With SortCollectionA(GetSubFoldersA(ScanCustomPath, GetSpecialFolderPathA(sfDesktop), , False, True))
      Do Until .Count = 0
        AddFolderItem .Item(1)
        .Remove 1
      Loop
    End With
  End If
  If m_ShowFiles Then
    With SortCollectionA(GetSubFilesA(ScanCustomPath, GetSpecialFolderPathA(sfDesktop), , False, True))
      Do Until .Count = 0
        AddFileItem .Item(1)
        .Remove 1
      Loop
    End With
  End If
End Sub

Private Sub AddMyComputerItems()
  Dim d As New PowerDLL2.Drives
  Dim i As Long
  SetColumnHeaderNames True
  For i = 1 To d.Count
    AddMyComputerItem d.Drive(i)
  Next
End Sub

Private Sub AddMyComputerItem(Drive As PowerDLL2.Drive)
  With m_ListViewParent.ListItems.Add(, (Drive.DriveLetter + ":"), GetDriveCaption(Drive))
    .SubItems(1) = GetDriveTypeName(Drive.DriveLetter)
    .SubItems(2) = GetByteStringA(Drive.TotalSpace)
    .SubItems(3) = GetByteStringA(Drive.FreeSpace)
    If UsingIcons Then
      .SmallIcon = GetIconIndex(Drive.Icon)
    End If
  End With
End Sub

Private Sub AddFolderItem(Path As String)
  With m_ListViewParent.ListItems.Add(, (FOLDER_PREFIX + Path), GetFileNameA(Path))
    .SubItems(1) = GetParentFolderName(Path)
    .SubItems(2) = GetByteStringA(0)
    .SubItems(3) = GetFileTypeDescriptionA(Path)
    .SubItems(4) = Trim(Str(GetFileDateA(Path)))
    If UsingIcons Then
      .SmallIcon = GetIconIndex(GetAssociatedIconA(Path))
    End If
  End With
End Sub

Private Sub AddFileItem(Path As String)
  With m_ListViewParent.ListItems.Add(, (FILE_PREFIX + Path), GetFileNameA(Path))
    .SubItems(1) = GetParentFolderName(Path)
    .SubItems(2) = GetByteStringA(FileLen(Path))
    .SubItems(3) = GetFileTypeDescriptionA(Path)
    .SubItems(4) = Trim(Str(GetFileDateA(Path)))
    If UsingIcons Then
      .SmallIcon = GetIconIndex(GetAssociatedIconA(Path))
    End If
  End With
End Sub

Private Function GetIconIndex(NewIcon As IPictureDisp) As Long
  If NewIcon Is Nothing Then
    Load frmObjects
    GetIconIndex = imlIcons.ListImages.Add(, , frmObjects.picNothing.Picture).Index
    Unload frmObjects
    Set frmObjects = Nothing
  Else
    GetIconIndex = imlIcons.ListImages.Add(, , NewIcon).Index
  End If
End Function

Private Function GetParentFolderName(Path As String) As String
  Dim pPath As String
  pPath = GetParentFolderA(Path)
  If LCase(pPath) = LCase(GetSpecialFolderPathA(sfDesktop)) Then
    GetParentFolderName = GetDesktopName
  Else
    GetParentFolderName = pPath
  End If
End Function

Private Function IsValidPath(Path As String) As Boolean
  IsValidPath = IsValidBrowserPath(Path)
End Function

Private Function GetSelectedItems() As String
  Dim sItem As MSComctlLib.ListItem
  Dim sItems As String
  If CanRun Then
    For Each sItem In m_ListViewParent.ListItems
      If sItem.Selected Then
        sItems = (sItems + (GetItemPath(sItem) + ";"))
      End If
    Next
    If Not sItems = "" Then
      sItems = Left(sItems, (Len(sItems) - 1))
      GetSelectedItems = sItems
    End If
  End If
End Function

Private Function IsMultipleSelected() As Boolean
  Dim sItem As MSComctlLib.ListItem
  Dim iSel As Boolean
  If CanRun Then
    For Each sItem In m_ListViewParent.ListItems
      If sItem.Selected Then
        If iSel Then
          IsMultipleSelected = True
          Exit For
        Else
          iSel = True
        End If
      End If
    Next
  End If
End Function

Private Function GetItemPath(Item As MSComctlLib.ListItem) As String
  If Item.Key = BROWSER_KEY_MYCOMPUTER Then
    GetItemPath = BROWSER_PATH_MYCOMPUTER
  ElseIf Len(Item.Key) = 2 Then
    GetItemPath = Item.Key
  Else
    Select Case Left(Item.Key, 2)
      Case FOLDER_PREFIX, FILE_PREFIX
        GetItemPath = Right(Item.Key, (Len(Item.Key) - 2))
    End Select
  End If
End Function

Private Sub UserChangePath(Path As String)
  m_Path = Path
  Refresh
  UpdateDirShow
  RaiseEvent PathChange(Path)
End Sub

Private Function IsMouseOverItem() As Boolean
  IsMouseOverItem = Not (GetItemFromMousePosition Is Nothing)
End Function

Private Function ItemIsFolder(Item As MSComctlLib.ListItem) As Boolean
  If Item.Key = BROWSER_KEY_MYCOMPUTER Then
    ItemIsFolder = True
  ElseIf Len(Item.Key) = 2 Then
    ItemIsFolder = True
  Else
    ItemIsFolder = (Left(Item.Key, 2) = FOLDER_PREFIX)
  End If
End Function

Private Function GetItemFromMousePosition() As MSComctlLib.ListItem
  Dim mPos As POINTAPI
  mPos = GetMousePositionA
  ScreenToClientA m_ListViewParent.hWnd, mPos
  mPos.x = (mPos.x * Screen.TwipsPerPixelX)
  mPos.y = (mPos.y * Screen.TwipsPerPixelY)
  Set GetItemFromMousePosition = m_ListViewParent.HitTest(mPos.x, mPos.y)
End Function

Private Sub ItemClick(Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    If ItemIsFolder(Item) Then
      RaiseEvent FolderClick(GetItemPath(Item))
    Else
      RaiseEvent FileClick(GetItemPath(Item))
    End If
  End If
  Set SelItem = Item
End Sub

Private Sub UserPressReturn()
  Dim sItems As String
  Dim eItem As String
  Dim eCancel As Boolean
  sItems = GetSelectedItems
  If Not sItems = "" Then
    If InStr(sItems, ";") = 0 Then
      If IsValidBrowserPath(sItems) Then
        UserChangePath sItems
      Else
        eCancel = False
        RaiseEvent ExecutingFile(sItems, eCancel)
        If Not eCancel Then
          ExecuteFileA sItems
        End If
      End If
    Else
      Do Until sItems = ""
        If InStr(sItems, ";") = 0 Then
          eItem = sItems
          sItems = ""
        Else
          eItem = Left(sItems, (InStr(sItems, ";") - 1))
          sItems = Right(sItems, (Len(sItems) - (Len(eItem) + 1)))
        End If
        eCancel = False
        RaiseEvent ExecutingFile(eItem, eCancel)
        If Not eCancel Then
          ExecuteFileA eItem
        End If
      Loop
    End If
  End If
End Sub

Private Sub UpdateDirShow()
  If Not m_DirShow Is Nothing Then
    m_DirShow.Path = m_Path
  End If
End Sub

Private Sub GoToParentFolder()
  If LCase(m_Path) = LCase(BROWSER_PATH_MYCOMPUTER) Then
    UserChangePath GetSpecialFolderPathA(sfDesktop)
  Else
    If Len(m_Path) = 2 Then
      UserChangePath BROWSER_PATH_MYCOMPUTER
    ElseIf Not LCase(m_Path) = LCase(GetSpecialFolderPathA(sfDesktop)) Then
      UserChangePath GetParentFolderA(m_Path)
    End If
  End If
End Sub

Private Function GetDeletableItems(SelItems As String) As String
  Dim sItems As String
  Dim nItem As String
  Do Until SelItems = ""
    If InStr(SelItems, ";") = 0 Then
      nItem = SelItems
      SelItems = ""
    Else
      nItem = Left(SelItems, (InStr(SelItems, ";") - 1))
      SelItems = Right(SelItems, (Len(SelItems) - (Len(nItem) + 1)))
    End If
    If CanDeletePath(nItem) Then
      sItems = (sItems + (nItem + ";"))
    End If
  Loop
  If Not sItems = "" Then
    GetDeletableItems = Left(sItems, (Len(sItems) - 1))
  End If
End Function

Private Sub DeleteSelectedItems(ShiftState As VBRUN.ShiftConstants)
  Dim sItems As String
  Dim dFlags As FILEOPERATION_FLAGS_Constants
  sItems = GetDeletableItems(GetSelectedItems)
  If Not ShiftState = vbShiftMask Then
    dFlags = fofAllowUndo
  End If
  ShellFileOperationA m_ListViewParent.hWnd, sItems, , foDelete, dFlags
  Refresh
End Sub

Private Sub Class_Initialize()
  m_Path = GetSpecialFolderPathA(sfDesktop)
  m_ShowFolders = True
  m_ShowFiles = True
  m_AllowRename = True
End Sub

Private Sub Class_Terminate()
  Detach
End Sub

Private Sub m_ListViewParent_AfterLabelEdit(Cancel As Integer, NewString As String)
  Dim oName As String
  Dim nName As String
  Dim iFile As Boolean
  If DiskRename Then
    GetDriveObjectFromDriveLetterA(Left(m_ListViewParent.SelectedItem.Key, 1)).VolumeLabel = NewString
    m_ListViewParent.SelectedItem.Text = GetDriveCaption(GetDriveObjectFromDriveLetterA(Left(m_ListViewParent.SelectedItem.Key, 1)))
  Else
    oName = m_ListViewParent.SelectedItem.Key
    iFile = Not ItemIsFolder(m_ListViewParent.SelectedItem)
    oName = Right(oName, (Len(oName) - 2))
    nName = ((GetParentFolderA(oName) + "\") + nName)
    If RenameFile(oName, nName) Then
      With m_ListViewParent.SelectedItem
        If iFile Then
          .Key = (FILE_PREFIX + nName)
        Else
          .Key = (FOLDER_PREFIX + nName)
        End If
      End With
      If iFile Then
        RaiseEvent FileRenamed(oName, nName)
      Else
        RaiseEvent FolderRenamed(oName, nName)
      End If
    Else
      Cancel = 1
    End If
  End If
End Sub

Private Sub m_ListViewParent_BeforeLabelEdit(Cancel As Integer)
  If SelectedItem = BROWSER_PATH_MYCOMPUTER Then
    Cancel = 1
  Else
    If Len(SelectedItem) = 2 Then
      m_ListViewParent.SelectedItem.Text = GetDriveObjectFromDriveLetterA(Left(SelectedItem, 1)).VolumeLabel
    End If
  End If
End Sub

Private Sub m_ListViewParent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  m_ListViewParent.Sorted = True
  m_ListViewParent.SortKey = ColumnHeader.Index
End Sub

Private Sub m_ListViewParent_DblClick()
  If Not InStr(GetSelectedItems, ";") = 0 Then
    UserPressReturn
  Else
    Dim mItem As MSComctlLib.ListItem
    Set mItem = GetItemFromMousePosition
    If Not mItem Is Nothing Then
      If ItemIsFolder(mItem) Then
        UserChangePath GetItemPath(mItem)
      Else
        RaiseEvent FileDblClick(GetItemPath(mItem))
      End If
    End If
  End If
End Sub

Private Sub m_ListViewParent_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF2
      If m_AllowRename Then
        m_ListViewParent.StartLabelEdit
      End If
    Case vbKeyF5
      Refresh
    Case vbKeyReturn
      UserPressReturn
    Case vbKeyBack
      GoToParentFolder
    Case vbKeyDelete
      DeleteSelectedItems GetShiftConstantFromCode(Shift)
  End Select
  If Not SelItem Is m_ListViewParent.SelectedItem Then
    ItemClick m_ListViewParent.SelectedItem
  End If
  RaiseEvent KeyDown(GetKeyConstantFromCode(KeyCode), GetShiftConstantFromCode(Shift))
End Sub

Private Sub m_ListViewParent_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not SelItem Is m_ListViewParent.SelectedItem Then
    ItemClick m_ListViewParent.SelectedItem
  End If
End Sub

Private Sub m_ListViewParent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    m_ListViewParent.DropHighlight = m_ListViewParent.HitTest(x, y)
  End If
End Sub

Private Sub m_ListViewParent_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim mItem As MSComctlLib.ListItem
  If Button = vbRightButton Then
    Set mItem = m_ListViewParent.HitTest(x, y)
    If Not mItem Is Nothing Then
      If ItemIsFolder(mItem) Then
        RaiseEvent FolderRightClick(GetItemPath(mItem))
      Else
        If Not mItem.Key = BROWSER_KEY_MYCOMPUTER Then
          RaiseEvent FileRightClick(GetItemPath(mItem))
        End If
      End If
    Else
      RaiseEvent RightClick
    End If
  ElseIf Button = vbLeftButton Then
    If Not SelItem Is m_ListViewParent.SelectedItem Then
      ItemClick m_ListViewParent.SelectedItem
    End If
  End If
End Sub

Private Sub m_ListViewParent_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim dFiles As String
  Dim dItem As MSComctlLib.ListItem
  Dim dPath As String
  dFiles = GetFilesFromDragDropData(Data)
  If Not dFiles = "" Then
    Set dItem = m_ListViewParent.HitTest(x, y)
    If Not dItem Is Nothing Then
      If ItemIsFolder(dItem) Then
        dPath = GetItemPath(dItem)
      Else
        Set dItem = Nothing
      End If
    End If
    If dItem Is Nothing Then
      dPath = m_Path
    End If
    If Not dPath = BROWSER_PATH_MYCOMPUTER Then
      ShellFileOperationA m_ListViewParent.hWnd, dFiles, dPath, foCopy, fofAllowUndo
      If dPath = m_Path Then
        Refresh
      End If
    End If
  End If
End Sub

Private Sub m_ListViewParent_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  m_ListViewParent.DropHighlight = m_ListViewParent.HitTest(x, y)
End Sub

Private Sub m_ListViewParent_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  Dim sItems As String
  Dim nItem As String
  AllowedEffects = ccOLEDropEffectNone
  sItems = GetSelectedItems
  If Not sItems = "" Then
    If InStr(sItems, ";") = 0 Then
      If sItems = BROWSER_PATH_MYCOMPUTER Then
        Exit Sub
      End If
    End If
    AllowedEffects = ccOLEDropEffectCopy
    Data.SetData , ccCFFiles
    Do Until sItems = ""
      If InStr(sItems, ";") = 0 Then
        nItem = sItems
        sItems = ""
      Else
        nItem = Left(sItems, (InStr(sItems, ";") - 1))
        sItems = Right(sItems, (Len(sItems) - (Len(nItem) + 1)))
      End If
      Data.Files.Add nItem
    Loop
  End If
End Sub
