VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMisc"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type VideoSize
  Width As Long
  Height As Long
End Type

Public Enum MULTIMEDIA_COMMAND
  mcOpen = &H1
  mcClose = &H2
  mcPlay = &H4
  mcPause = &H8
  mcStop = &H10
  mcRewind = &H20
End Enum

Public Enum COMMONCONTROLS_CLASSES
  ccListView_Classes = ICC_LISTVIEW_CLASSES
  ccTreeView_Classes = ICC_TREEVIEW_CLASSES
  ccToolBar_Classes = ICC_BAR_CLASSES
  ccTab_Classes = ICC_TAB_CLASSES
  ccUpDown_Classes = ICC_UPDOWN_CLASS
  ccProgress_Class = ICC_PROGRESS_CLASS
  ccHotkey_Class = ICC_HOTKEY_CLASS
  ccAnimate_Class = ICC_ANIMATE_CLASS
  ccWin95_Classes = ICC_WIN95_CLASSES
  ccCalendar_Classes = ICC_DATE_CLASSES
  ccComboEx_Classes = ICC_USEREX_CLASSES
  ccCoolBar_Classes = ICC_COOL_CLASSES
  ccInternet_Classes = ICC_INTERNET_CLASSES
  ccPageScroller_Class = ICC_PAGESCROLLER_CLASS
  ccNativeFont_Class = ICC_NATIVEFNTCTL_CLASS
  ccStandard_Classes = ICC_STANDARD_CLASSES
  ccLink_Class = ICC_LINK_CLASS
  ccAll_Classes = ccListView_Classes Or ccTreeView_Classes Or ccToolBar_Classes Or ccTab_Classes Or ccUpDown_Classes Or ccProgress_Class Or ccHotkey_Class Or ccAnimate_Class Or ccWin95_Classes Or ccCalendar_Classes Or ccComboEx_Classes Or ccCoolBar_Classes Or ccInternet_Classes Or ccPageScroller_Class Or ccNativeFont_Class Or ccStandard_Classes Or ccLink_Class
End Enum

Public Function InitCommonControls(Optional ccFlags As COMMONCONTROLS_CLASSES = ccAll_Classes) As Boolean
  InitCommonControls = InitCommonControlsA(ccFlags)
End Function

Public Function GetFileName(Path As String) As String
  GetFileName = GetFileNameA(Path)
End Function

Public Function GetParentFolder(Path As String) As String
  GetParentFolder = GetParentFolderA(Path)
End Function

Public Function GetFileExtension(Path As String) As String
  GetFileExtension = GetFileExtensionA(Path)
End Function

Public Function GetRandomNumber(Min As Double, Max As Double) As Double
  GetRandomNumber = GetRandomNumberA(Min, Max)
End Function

Public Function FixPath(Path As String) As String
  FixPath = FixPathA(Path)
End Function

Public Function QuotePath(Path As String) As String
  QuotePath = QuotePathA(Path)
End Function

Public Function UnQuotePath(Path As String) As String
  UnQuotePath = UnQuotePathA(Path)
End Function

Public Sub AddCollectionToCollection(SourceCol As Collection, AddCol As Collection)
  AddCollectionToCollectionA SourceCol, AddCol
End Sub

Public Function GetByteString(Bytes As Double) As String
  GetByteString = GetByteStringA(Bytes)
End Function

Public Sub CenterForm(Form As Object)
  CenterFormA Form
End Sub

Public Sub AutoSizeListviewColumns(ListviewObject As Object)
  AutoSizeListviewColumnsA ListviewObject
End Sub

Public Sub SplitRGB(Color As Long, Red As Integer, Green As Integer, Blue As Integer)
  SplitRGBA Color, Red, Green, Blue
End Sub

Public Function IsInternetOffline() As Boolean
  IsInternetOffline = IsInternetOfflineA
End Function

Public Function IsInternetOnline() As Boolean
  IsInternetOnline = IsInternetOnlineA
End Function

Public Function GetCurrentDisplayMode() As clsDisplayMode
  Set GetCurrentDisplayMode = GetCurrentDisplayModeA
End Function
