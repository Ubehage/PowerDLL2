Attribute VB_Name = "modSorting"
Option Explicit

Public Sub SortStringArrayA(sArray() As String, Min As Long, Max As Long)
  Dim i As Long
  Dim j As Long
  Dim mString As String
  Dim tString As String
  If Max > Min Then
    mString = LCase(sArray(((Min + Max) / 2)))
    i = Min
    j = Max
    Do While i <= j
      If (LCase(sArray(i)) >= mString And LCase(sArray(j)) <= mString) Then
        tString = sArray(i)
        sArray(i) = sArray(j)
        sArray(j) = tString
        i = (i + 1)
        j = (j - 1)
      Else
        If LCase(sArray(i)) < mString Then
          i = (i + 1)
        End If
        If LCase(sArray(j)) > mString Then
          j = (j - 1)
        End If
      End If
    Loop
    SortStringArrayA sArray(), Min, j
    SortStringArrayA sArray(), i, Max
  End If
End Sub

Public Function SortCollectionA(SortCol As Collection) As Collection
  Dim i As Long
  Dim sArray() As String
  Set SortCollectionA = New Collection
  If Not SortCol.Count = 0 Then
    ReDim sArray(1 To SortCol.Count) As String
    For i = LBound(sArray) To UBound(sArray)
      sArray(i) = SortCol.Item(i)
    Next
    SortStringArrayA sArray(), LBound(sArray), UBound(sArray)
    For i = LBound(sArray) To UBound(sArray)
      SortCollectionA.Add sArray(i)
    Next
    Erase sArray
  End If
End Function

Public Function RandomizeCollectionA(RandCol As Collection) As Collection
  Dim i As Long
  Dim tCol As Collection
  Set tCol = New Collection
  For i = 1 To RandCol.Count
    tCol.Add RandCol.Item(i)
  Next
  Set RandomizeCollectionA = New Collection
  Do Until tCol.Count = 0
    i = GetRandomNumberA(1, tCol.Count)
    RandomizeCollectionA.Add tCol.Item(i)
    tCol.Remove i
  Loop
  Set tCol = Nothing
End Function
