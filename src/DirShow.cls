VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents pTimer As PowerDLL2.Timer
Attribute pTimer.VB_VarHelpID = -1

Dim WithEvents dWalk As PowerDLL2.DirWalk
Attribute dWalk.VB_VarHelpID = -1

Dim WithEvents m_TreeViewParent As MSComctlLib.TreeView
Attribute m_TreeViewParent.VB_VarHelpID = -1
Dim m_AllowRename As Boolean
Dim m_Path As String
Dim m_FileShow As PowerDLL2.FileShow

Dim imlIcons As MSComctlLib.ImageList

Dim HasFolders As Boolean

Dim NamingDisk As Boolean

Dim ChangedByCode As Boolean

Dim ClickedNode As MSComctlLib.Node

Public Event PathChange(Path As String)
Public Event FolderRenamed(OldPath As String, NewPath As String)
Public Event FolderRightClick(Path As String)
Public Event KeyDown(KeyCode As VBRUN.KeyCodeConstants, Shift As VBRUN.ShiftConstants)

Public Property Get TreeViewParent() As Object
  Set TreeViewParent = m_TreeViewParent
End Property
Public Property Get AllowRename() As Boolean
  AllowRename = m_AllowRename
End Property
Public Property Get Path() As String
  Path = m_Path
End Property
Public Property Get FileShow() As PowerDLL2.FileShow
  Set FileShow = m_FileShow
End Property
Public Property Let TreeViewParent(New_TreeViewParent As Object)
  If Not m_TreeViewParent Is New_TreeViewParent Then
    Set m_TreeViewParent = New_TreeViewParent
    If m_TreeViewParent Is Nothing Then
      Set pTimer = Nothing
    Else
      If pTimer Is Nothing Then
        Set pTimer = New PowerDLL2.Timer
      End If
      pTimer.hWndParent = m_TreeViewParent.hWnd
      pTimer.Interval = 300
      pTimer.Enabled = False
    End If
    InitTreeView
    SetAllowRename
    Refresh
  End If
End Property
Public Property Let AllowRename(New_AllowRename As Boolean)
  If Not m_AllowRename = New_AllowRename Then
    m_AllowRename = New_AllowRename
    SetAllowRename
  End If
End Property
Public Property Let Path(New_Path As String)
  If Not m_Path = New_Path Then
    If IsValidBrowserPath(New_Path) Then
      m_Path = New_Path
      SelectPath m_Path
      ChangePath m_Path
    End If
  End If
End Property
Public Property Let FileShow(New_FileShow As PowerDLL2.FileShow)
  If Not m_FileShow Is New_FileShow Then
    Set m_FileShow = New_FileShow
    If Not m_FileShow Is Nothing Then
      m_FileShow.DirShow = Me
    End If
    UpdateFileShow
  End If
End Property

Public Sub Refresh()
  If Not m_TreeViewParent Is Nothing Then
    ClearTreeView
    AddDesktopNode
    If Not IsFolderA(m_Path) Then
      If Not LCase(m_Path) = LCase(BROWSER_PATH_MYCOMPUTER) Then
        m_Path = GetSpecialFolderPathA(sfDesktop)
      End If
    End If
    SelectPath m_Path
    UpdateFileShow
  End If
End Sub

Public Sub Detach()
  If Not m_TreeViewParent Is Nothing Then
    ClearTreeView
  End If
End Sub

Public Sub GoToPath(NewPath As String)
  m_Path = GetPathFromBrowserString(m_Path, NewPath)
  Refresh
  'If IsValidBrowserPath(NewPath) Then
  '  Path = NewPath
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

Private Sub InitTreeView()
  With m_TreeViewParent
    .Indentation = 10
    .OLEDragMode = ccOLEDragAutomatic
    .OLEDropMode = ccOLEDropManual
  End With
End Sub

Private Sub SetAllowRename()
  If Not m_TreeViewParent Is Nothing Then
    If m_AllowRename Then
      m_TreeViewParent.LabelEdit = tvwAutomatic
    Else
      m_TreeViewParent.LabelEdit = tvwManual
    End If
  End If
End Sub

Private Sub ClearTreeNodes()
  If Not m_TreeViewParent Is Nothing Then
    m_TreeViewParent.Nodes.Clear
  End If
End Sub

Private Sub ClearImageList()
  If Not m_TreeViewParent Is Nothing Then
    If m_TreeViewParent.ImageList Is Nothing Then
      Set imlIcons = Nothing
    Else
      Set imlIcons = m_TreeViewParent.ImageList
      Set m_TreeViewParent.ImageList = Nothing
      imlIcons.ListImages.Clear
      imlIcons.ImageWidth = 16
      imlIcons.ImageHeight = 16
      imlIcons.ListImages.Add , , LoadResPicture(101, vbResIcon)
      Set m_TreeViewParent.ImageList = imlIcons
    End If
  End If
End Sub

Private Sub ClearTreeView()
  ClearTreeNodes
  ClearImageList
End Sub

Private Sub AddDesktopNode()
  If UsingIcons Then
    ExpandNode m_TreeViewParent.Nodes.Add(, , BROWSER_KEY_DESKTOP, GetDesktopName, GetIconIndex(GetDesktopIcon))
  Else
    ExpandNode m_TreeViewParent.Nodes.Add(, , BROWSER_KEY_DESKTOP, GetDesktopName)
  End If
End Sub

Private Sub AddDesktopSubNodes()
  If UsingIcons Then
    AddNodeExpander m_TreeViewParent.Nodes.Add(BROWSER_KEY_DESKTOP, tvwChild, BROWSER_KEY_DOCUMENTS, GetDocumentsName, GetIconIndex(GetMyDocumentsIcon))
    AddNodeExpander m_TreeViewParent.Nodes.Add(BROWSER_KEY_DESKTOP, tvwChild, BROWSER_KEY_MYCOMPUTER, GetMyComputerName, GetIconIndex(GetMyComputerIcon))
  Else
    AddNodeExpander m_TreeViewParent.Nodes.Add(BROWSER_KEY_DESKTOP, tvwChild, BROWSER_KEY_DOCUMENTS, GetDocumentsName)
    AddNodeExpander m_TreeViewParent.Nodes.Add(BROWSER_KEY_DESKTOP, tvwChild, BROWSER_KEY_MYCOMPUTER, GetMyComputerName)
  End If
End Sub

Private Sub AddMyComputerSubNodes()
  Dim d As New PowerDLL2.Drives
  Dim i As Long
  For i = 1 To d.Count
    With d.Drive(i)
      If UsingIcons Then
        AddNodeExpander m_TreeViewParent.Nodes.Add(BROWSER_KEY_MYCOMPUTER, tvwChild, (.DriveLetter + ":"), GetDriveCaption(d.Drive(i)), GetIconIndex(.Icon))
      Else
        AddNodeExpander m_TreeViewParent.Nodes.Add(BROWSER_KEY_MYCOMPUTER, tvwChild, (.DriveLetter + ":"), GetDriveCaption(d.Drive(i)))
      End If
    End With
  Next
End Sub

Private Function GetIconIndex(NewIcon As IPictureDisp) As Long
  If UsingIcons Then
    If NewIcon Is Nothing Then
      Load frmObjects
      GetIconIndex = imlIcons.ListImages.Add(, , frmObjects.picNothing.Picture).Index
      Unload frmObjects
      Set frmObjects = Nothing
    Else
      GetIconIndex = imlIcons.ListImages.Add(, , NewIcon).Index
    End If
  End If
End Function

Private Function UsingIcons() As Boolean
  UsingIcons = Not (imlIcons Is Nothing)
End Function

Private Sub AddNodeExpander(pNode As MSComctlLib.Node)
  If HasSubNodes(pNode) Then
    m_TreeViewParent.Nodes.Add pNode.Key, tvwChild, (pNode.Key + "\expander")
  End If
End Sub

Private Sub RemoveSubNodes(pNode As MSComctlLib.Node)
  While pNode.Children > 0
    m_TreeViewParent.Nodes.Remove pNode.Child.Index
  Wend
End Sub

Private Sub AddSubNodes(pNode As MSComctlLib.Node)
  Select Case pNode.Key
    Case BROWSER_KEY_DESKTOP
      AddDesktopSubNodes
    Case BROWSER_KEY_MYCOMPUTER
      AddMyComputerSubNodes
    Case Else
      With SortCollectionA(GetSubFoldersA(ScanCustomPath, GetNodePath(pNode), , False, True))
        Do Until .Count = 0
          If UsingIcons Then
            AddNodeExpander m_TreeViewParent.Nodes.Add(pNode.Key, tvwChild, ((pNode.Key + "\") + GetFileNameA(.Item(1))), GetFileNameA(.Item(1)), GetIconIndex(GetAssociatedIconA(.Item(1))))
          Else
            AddNodeExpander m_TreeViewParent.Nodes.Add(pNode.Key, tvwChild, ((pNode.Key + "\") + GetFileNameA(.Item(1))), GetFileNameA(.Item(1)))
          End If
          .Remove 1
        Loop
      End With
  End Select
End Sub

Private Sub CollapseNode(pNode As MSComctlLib.Node)
  ChangedByCode = True
  pNode.Expanded = False
  ChangedByCode = False
End Sub

Private Sub ExpandNode(pNode As MSComctlLib.Node)
  RemoveSubNodes pNode
  ChangedByCode = True
  pNode.Expanded = True
  ChangedByCode = False
  AddSubNodes pNode
End Sub

Private Function HasSubNodes(pNode As MSComctlLib.Node) As Boolean
  If Len(pNode.Key) = 2 Then
    HasSubNodes = True
  Else
    Select Case pNode.Key
      Case BROWSER_KEY_DESKTOP
        HasSubNodes = True
      Case BROWSER_KEY_MYCOMPUTER
        HasSubNodes = True
      Case Else
        HasFolders = False
        dWalk.SearchPath GetNodePath(pNode), , False, True, False, True
        HasSubNodes = HasFolders
    End Select
  End If
End Function

Private Function GetNodePath(pNode As MSComctlLib.Node) As String
  Dim nPath As String
  nPath = pNode.Key
  nPath = Replace(nPath, BROWSER_KEY_DESKTOP, GetSpecialFolderPathA(sfDesktop))
  nPath = Replace(nPath, BROWSER_KEY_DOCUMENTS, GetSpecialFolderPathA(sfDocuments))
  nPath = Replace(nPath, BROWSER_KEY_MYCOMPUTER, BROWSER_PATH_MYCOMPUTER)
  GetNodePath = nPath
End Function

Private Sub SelectPath(Path As String)
  If Not m_TreeViewParent Is Nothing Then
    With TrackPath(Path)
      .Selected = True
      .EnsureVisible
    End With
  End If
End Sub

Private Function TrackPath(Path As String) As MSComctlLib.Node
  Dim tPath As String
  tPath = Replace2(Path, GetSpecialFolderPathA(sfDesktop), BROWSER_KEY_DESKTOP)
  tPath = Replace2(tPath, GetSpecialFolderPathA(sfDocuments), BROWSER_KEY_DOCUMENTS)
  tPath = Replace2(tPath, BROWSER_PATH_MYCOMPUTER, BROWSER_KEY_MYCOMPUTER)
  If Mid(tPath, 2, 1) = ":" Then
    ExpandNode GetNodeFromKey(BROWSER_KEY_MYCOMPUTER)
  End If
  With GetFoldersFromPath(tPath)
    tPath = ""
    Do Until .Count = 1
      If tPath = "" Then
        tPath = .Item(1)
      Else
        tPath = ((tPath + "\") + .Item(1))
      End If
      ExpandNode GetNodeFromKey(tPath)
      .Remove 1
    Loop
  End With
  Set TrackPath = GetNodeFromPath(Path)
End Function

Private Function GetNodeFromPath(nPath As String) As MSComctlLib.Node
  Set GetNodeFromPath = GetNodeFromKey(Replace2(Replace2(Replace2(nPath, GetSpecialFolderPathA(sfDesktop), BROWSER_KEY_DESKTOP), GetSpecialFolderPathA(sfDocuments), BROWSER_KEY_DOCUMENTS), BROWSER_PATH_MYCOMPUTER, BROWSER_KEY_MYCOMPUTER))
End Function

Private Function GetNodeFromKey(nKey As String) As MSComctlLib.Node
  Dim nItem As MSComctlLib.Node
  If Not m_TreeViewParent Is Nothing Then
    For Each nItem In m_TreeViewParent.Nodes
      If LCase(nKey) = LCase(nItem.Key) Then
        Set GetNodeFromKey = nItem
        Exit For
      End If
    Next
  End If
End Function

Private Function GetNodeFromMouse() As MSComctlLib.Node
  Dim mPos As POINTAPI
  mPos = GetMousePositionA
  ScreenToClientA m_TreeViewParent.hWnd, mPos
  With mPos
    .x = (.x * Screen.TwipsPerPixelX)
    .y = (.y * Screen.TwipsPerPixelY)
    Set GetNodeFromMouse = m_TreeViewParent.HitTest(.x, .y)
  End With
End Function

Private Sub ChangePath(NewPath As String)
  If Not m_Path = NewPath Then
    m_Path = NewPath
  End If
  UpdateFileShow
  RaiseEvent PathChange(NewPath)
End Sub

Private Sub UpdateFileShow()
  If Not m_FileShow Is Nothing Then
    pTimer.RestartTimer
  End If
End Sub

Private Sub DeleteSelectedItem(ShiftState As VBRUN.ShiftConstants)
  Dim dPath As String
  Dim dFlags As FILEOPERATION_FLAGS_Constants
  dPath = GetNodePath(m_TreeViewParent.SelectedItem)
  If CanDeletePath(dPath) Then
    If Not ShiftState = vbShiftMask Then
      dFlags = fofAllowUndo
    End If
    ShellFileOperationA m_TreeViewParent.hWnd, dPath, , foDelete, dFlags
    If Not IsValidBrowserPath(GetNodePath(m_TreeViewParent.SelectedItem)) Then
      m_Path = GetNodePath(m_TreeViewParent.SelectedItem.Parent)
      UpdateFileShow
    End If
    Refresh
  End If
End Sub

Private Sub dWalk_ItemFound(Path As String, Name As String, IsFile As Boolean, Cancel As Boolean)
  HasFolders = True
  Cancel = True
End Sub

Private Sub m_TreeViewParent_AfterLabelEdit(Cancel As Integer, NewString As String)
  With m_TreeViewParent.SelectedItem
    If NamingDisk Then
      GetDriveObjectFromDriveLetterA(Left(.Key, 1)).VolumeLabel = NewString
      .Text = GetDriveCaption(GetDriveObjectFromDriveLetterA(Left(.Key, 1)))
    Else
      If RenameFolder(GetNodePath(m_TreeViewParent.SelectedItem), NewString) Then
        RaiseEvent FolderRenamed(GetNodePath(m_TreeViewParent.SelectedItem), ((GetParentFolderA(GetNodePath(m_TreeViewParent.SelectedItem)) + "\") + NewString))
      Else
        Cancel = 1
      End If
    End If
  End With
End Sub

Private Sub m_TreeViewParent_BeforeLabelEdit(Cancel As Integer)
  Select Case m_TreeViewParent.SelectedItem.Key
    Case BROWSER_KEY_DESKTOP, BROWSER_KEY_DOCUMENTS, BROWSER_KEY_MYCOMPUTER
      Cancel = 1
    Case Else
      If Len(m_TreeViewParent.SelectedItem.Key) = 2 Then
        m_TreeViewParent.SelectedItem.Text = GetDriveObjectFromDriveLetterA(Left(m_TreeViewParent.SelectedItem.Key, 1)).VolumeLabel
        NamingDisk = True
      Else
        NamingDisk = False
      End If
  End Select
End Sub

Private Sub m_TreeViewParent_Collapse(ByVal Node As MSComctlLib.Node)
  If Not ChangedByCode Then
    CollapseNode Node
  End If
End Sub

Private Sub m_TreeViewParent_Expand(ByVal Node As MSComctlLib.Node)
  If Not ChangedByCode Then
    ExpandNode Node
  End If
End Sub

Private Sub m_TreeViewParent_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF2
      If m_AllowRename Then
        m_TreeViewParent.StartLabelEdit
      End If
    Case vbKeyF5
      Refresh
    Case vbKeyDelete
      DeleteSelectedItem GetShiftConstantFromCode(Shift)
  End Select
  RaiseEvent KeyDown(GetKeyConstantFromCode(KeyCode), GetShiftConstantFromCode(Shift))
End Sub

Private Sub m_TreeViewParent_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not GetNodePath(m_TreeViewParent.SelectedItem) = m_Path Then
    ChangePath GetNodePath(m_TreeViewParent.SelectedItem)
  End If
End Sub

Private Sub m_TreeViewParent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    If Not m_TreeViewParent.HitTest(x, y) Is Nothing Then
      m_TreeViewParent.DropHighlight = m_TreeViewParent.HitTest(x, y)
    End If
  ElseIf Button = vbLeftButton Then
    Set ClickedNode = m_TreeViewParent.HitTest(x, y)
  End If
End Sub

Private Sub m_TreeViewParent_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    If Not m_TreeViewParent.HitTest(x, y) Is Nothing Then
      RaiseEvent FolderRightClick(GetNodePath(m_TreeViewParent.HitTest(x, y)))
    End If
    m_TreeViewParent.DropHighlight = Nothing
  End If
End Sub

Private Sub m_TreeViewParent_NodeClick(ByVal Node As MSComctlLib.Node)
  ChangePath GetNodePath(Node)
End Sub

Private Sub m_TreeViewParent_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim dFiles As String
  Dim dNode As MSComctlLib.Node
  dFiles = GetFilesFromDragDropData(Data)
  If Not dFiles = "" Then
    Set dNode = m_TreeViewParent.HitTest(x, y)
    If dNode Is Nothing Then
      Set dNode = m_TreeViewParent.SelectedItem
    End If
    If Not dNode.Key = BROWSER_KEY_MYCOMPUTER Then
      ShellFileOperationA m_TreeViewParent.hWnd, dFiles, GetNodePath(dNode), foCopy, fofAllowUndo
      Refresh
    End If
  End If
End Sub

Private Sub m_TreeViewParent_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  If Not m_TreeViewParent.HitTest(x, y) Is Nothing Then
    m_TreeViewParent.DropHighlight = m_TreeViewParent.HitTest(x, y)
  Else
    m_TreeViewParent.DropHighlight = Nothing
  End If
End Sub

Private Sub m_TreeViewParent_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  Dim dNode As MSComctlLib.Node
  AllowedEffects = ccOLEDropEffectNone
  Set dNode = GetNodeFromMouse
  If dNode Is Nothing Then
    If ClickedNode Is Nothing Then
      If Not m_TreeViewParent.SelectedItem Is Nothing Then
        Set dNode = m_TreeViewParent.SelectedItem
      End If
    Else
      Set dNode = ClickedNode
    End If
  End If
  If Not dNode.Key = BROWSER_KEY_MYCOMPUTER Then
    AllowedEffects = ccOLEDropEffectCopy
    Data.SetData , ccCFFiles
    Data.Files.Add GetNodePath(dNode)
  End If
End Sub

Private Sub Class_Initialize()
  Set dWalk = New PowerDLL2.DirWalk
  m_Path = GetSpecialFolderPathA(sfDesktop)
  m_AllowRename = True
End Sub

Private Sub Class_Terminate()
  If Not pTimer Is Nothing Then
    Set pTimer = Nothing
  End If
  Set dWalk = Nothing
  Detach
End Sub

Private Sub pTimer_Timer()
  If Not m_FileShow Is Nothing Then
    m_FileShow.Path = m_Path
  End If
  pTimer.Enabled = False
End Sub
