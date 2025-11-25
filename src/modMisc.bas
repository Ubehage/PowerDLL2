Attribute VB_Name = "modMisc"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, Param As Any) As Long
Private Declare Sub InitCommonControls9x Lib "comctl32" Alias "InitCommonControls" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Type tagINITCOMMONCONTROLSEX
  dwSize As Long
  dwICC As Long
End Type

Public Function InitCommonControlsA(Optional ccFlags As COMMONCONTROLS_CLASSES = ccAll_Classes) As Boolean
  Dim icc As tagINITCOMMONCONTROLSEX
  On Error GoTo OldCC
  With icc
    .dwSize = Len(icc)
    .dwICC = ccFlags
  End With
  InitCommonControlsA = InitCommonControlsEx(icc)
ExitNow:
  On Error GoTo 0
  Exit Function
OldCC:
  InitCommonControls9x
  Resume ExitNow
End Function

Public Function GetFileNameA(Path As String) As String
  Dim i As Long
  Dim NameSet As Boolean
  For i = Len(Path) To 1 Step -1
    If Mid(Path, i, 1) = "\" Then
      GetFileNameA = Right(Path, (Len(Path) - i))
      NameSet = True
      Exit For
    End If
  Next
  If Not NameSet Then
    GetFileNameA = Path
  End If
End Function

Public Function GetParentFolderA(Path As String) As String
  Dim fName As String
  fName = GetFileNameA(Path)
  If Not fName = "" Then
    If Not fName = Path Then
      GetParentFolderA = Left(Path, (Len(Path) - (Len(fName) + 1)))
    End If
  End If
End Function

Public Function GetFileExtensionA(Path As String) As String
  Dim i As Long
  For i = Len(Path) To 1 Step -1
    If Mid(Path, i, 1) = "." Then
      GetFileExtensionA = Right(Path, (Len(Path) - i))
      Exit For
    End If
  Next
End Function

Public Function GetRandomNumberA(Min As Double, Max As Double) As Double
  Randomize Timer
  GetRandomNumberA = ((Rnd * Max) + Min)
End Function

Public Function FixPathA(Path As String) As String
  If Right(Path, 1) = "\" Then
    FixPathA = Left(Path, (Len(Path) - 1))
  Else
    FixPathA = Path
  End If
End Function

Public Function QuotePathA(Path As String) As String
  Dim qPath As String
  qPath = Path
  If Not Left(qPath, 1) = Chr(34) Then
    qPath = (Chr(34) + qPath)
  End If
  If Not Right(qPath, 1) = Chr(34) Then
    qPath = (qPath + Chr(34))
  End If
  QuotePathA = qPath
End Function

Public Function UnQuotePathA(Path As String) As String
  Dim qPath As String
  qPath = Path
  If Left(qPath, 1) = Chr(34) Then
    qPath = Right(qPath, (Len(qPath) - 1))
  End If
  If Right(qPath, 1) = Chr(34) Then
    qPath = Left(qPath, (Len(qPath) - 1))
  End If
  UnQuotePathA = qPath
End Function

Public Sub AddCollectionToCollectionA(SourceCol As Collection, AddCol As Collection)
  Dim i As Long
  If Not SourceCol Is Nothing Then
    If Not AddCol Is Nothing Then
      For i = 1 To AddCol.Count
        SourceCol.Add AddCol.Item(i)
      Next
    End If
  End If
End Sub

Public Function GetByteStringA(Bytes As Double) As String
  Dim bValue As Double
  bValue = Bytes
  If bValue < 1024 Then
    GetByteStringA = (SetMaxDecimals(bValue, 2) + " bytes")
  Else
    bValue = (bValue / 1024)
    If bValue < 1024 Then
      GetByteStringA = (SetMaxDecimals(bValue, 2) + " Kbytes")
    Else
      bValue = (bValue / 1024)
      If bValue < 1024 Then
        GetByteStringA = (SetMaxDecimals(bValue, 2) + " Mbytes")
      Else
        bValue = (bValue / 1024)
        If bValue < 1024 Then
          GetByteStringA = (SetMaxDecimals(bValue, 2) + " Gbytes")
        Else
          bValue = (bValue / 1024)
          If bValue < 1024 Then
            GetByteStringA = (SetMaxDecimals(bValue, 2) + " Tbytes")
          Else
            bValue = (bValue / 1024)
            GetByteStringA = (SetMaxDecimals(bValue, 2) + " Ebytes")
          End If
        End If
      End If
    End If
  End If
End Function

Public Sub CenterFormA(Form As Object)
  On Error Resume Next
  Form.Move ((Screen.Width - Form.Width) / 2), ((Screen.Height - Form.Height) / 2)
  On Error GoTo 0
End Sub

Public Sub AutoSizeListviewColumnsA(ListviewObject As Object)
  Dim i As Long
  On Error GoTo Error1
  For i = 0 To (ListviewObject.ColumnHeaders.Count - 1)
    SendMessage ListviewObject.hWnd, i, ByVal LVSCW_AUTOSIZE_USEHEADER, 0&
  Next
ExitNow:
  On Error GoTo 0
  Exit Sub
Error1:
  Resume ExitNow
End Sub

Public Sub SplitRGBA(Color As Long, Red As Integer, Green As Integer, Blue As Integer)
  Blue = (Color \ 65536) And &HFF
  Green = (Color \ 256) And &HFF
  Red = Color And &HFF
  'Red = (Color Mod 256)
  'Blue = Int((Color \ 65536))
  'Green = ((Color - (Blue * 65536) - Red) \ 256)
End Sub

Public Function IsInternetOfflineA() As Boolean
  IsInternetOfflineA = InetIsOffline(0)
End Function

Public Function IsInternetOnlineA() As Boolean
  IsInternetOnlineA = Not IsInternetOfflineA
End Function

Public Function GetCurrentDisplayModeA() As clsDisplayMode
  Dim dModes As clsDisplayModes
  Set dModes = New clsDisplayModes
  Set GetCurrentDisplayModeA = dModes.GetCurrentDisplayMode
  Set dModes = Nothing
End Function
