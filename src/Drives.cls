VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Count As Long
Dim m_Item() As Drive

Public Property Get Count() As Long
  Count = m_Count
End Property
Public Property Get Drive(Index As Variant) As Drive
  If IsNumeric(Index) Then
    Set Drive = m_Item(CLng(Index))
  Else
    Set Drive = m_Item(GetDriveIndex(LCase(Left(CStr(Index), 1))))
  End If
End Property

Public Sub Refresh()
  Clear
  FindAllDrives
End Sub

Private Function GetDriveIndex(Drive As String) As Long
  Dim i As Long
  For i = 1 To m_Count
    If Drive = LCase(m_Item(i).DriveLetter) Then
      GetDriveIndex = i
      Exit For
    End If
  Next
End Function

Private Sub FindAllDrives()
  Dim dLetters As String
  dLetters = GetDriveLettersA
  Do Until dLetters = ""
    AddNewDrive Left(dLetters, 1)
    dLetters = Right(dLetters, (Len(dLetters) - 1))
  Loop
End Sub

Private Sub AddNewDrive(DriveLetter As String)
  If (m_Count Mod 10) = 0 Then
    ReDim Preserve m_Item(1 To (m_Count + 10)) As Drive
  End If
  m_Count = (m_Count + 1)
  Set m_Item(m_Count) = New Drive
  m_Item(m_Count).DriveLetter = DriveLetter
End Sub

Private Sub Clear()
  m_Count = 0
  Erase m_Item
End Sub

Private Sub Class_Initialize()
  Refresh
End Sub

Private Sub Class_Terminate()
  Clear
End Sub
