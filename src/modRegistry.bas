Attribute VB_Name = "modRegistry"
Option Explicit

Public Function GetRegistryKeyA(hKey As REGISTRY_HKEYS, Optional Key As String, Optional ForceCreate As Boolean = False) As RegKey
  Dim pKey As RegKey
  Dim rKey As RegKey
  Set pKey = RegKeyFromHKey(hKey)
  If Key = "" Then
    Set rKey = pKey
  Else
    On Error Resume Next
    Set rKey = pKey.ParseKeyName(Key)
    On Error GoTo 0
    If rKey Is Nothing Then
      If ForceCreate Then
        Set rKey = CreateRegistryKeyA(hKey, Key)
      End If
    End If
  End If
  Set GetRegistryKeyA = rKey
End Function

Public Function CreateRegistryKeyA(hKey As REGISTRY_HKEYS, Key As String) As RegKey
  Dim nKey As String
  Dim pKey As RegKey
  Dim rKey As RegKey
  Dim rKeys As Collection
  Set rKeys = New Collection
  nKey = Key
  While Not InStr(nKey, "\") = 0
    rKeys.Add GetFileNameA(nKey)
    nKey = GetParentFolderA(nKey)
  Wend
  If Not nKey = "" Then
    rKeys.Add nKey
  End If
  Set pKey = RegKeyFromHKey(hKey)
  On Error Resume Next
  Do Until rKeys.Count = 0
    Set rKey = pKey.ParseKeyName(rKeys.Item(rKeys.Count))
    If rKey Is Nothing Then
      pKey.SubKeys.Add rKeys.Item(rKeys.Count)
      Set rKey = pKey.ParseKeyName(rKeys.Item(rKeys.Count))
    End If
    Set pKey = rKey
    Set rKey = Nothing
    If pKey Is Nothing Then
      Exit Do
    End If
    rKeys.Remove rKeys.Count
  Loop
  Set CreateRegistryKeyA = pKey
End Function

Public Function GetRegistryValueA(hKey As REGISTRY_HKEYS, Key As String, Optional ValueName As String) As Variant
  Dim rKey As RegKey
  Dim rValue As RegValue
  If ValueName = "" Then
    Set rKey = GetRegistryKeyA(hKey, Key)
    If Not rKey Is Nothing Then
      GetRegistryValueA = rKey.Value
    End If
  Else
    Set rValue = GetRegistryValueObjectA(hKey, Key, ValueName)
    If Not rValue Is Nothing Then
      GetRegistryValueA = rValue.Value
    End If
  End If
End Function

Public Function SetRegistryValueA(NewValue As Variant, hKey As REGISTRY_HKEYS, Key As String, Optional ValueName As String, Optional ForceCreate As Boolean) As Boolean
  Dim rKey As RegKey
  Dim rValue As RegValue
  Set rKey = GetRegistryKeyA(hKey, Key, True)
  If Not rKey Is Nothing Then
    Set rValue = rKey.Values.Item(ValueName)
    If rValue Is Nothing Then
      rKey.Values.Add ValueName
      Set rValue = rKey.Values.Item(ValueName)
    End If
    If Not rValue Is Nothing Then
      rValue.Value = NewValue
      SetRegistryValueA = True
    End If
  End If
End Function

Public Function DeleteRegistryValueA(hKey As REGISTRY_HKEYS, Key As String, ValueName As String) As Boolean
  Dim rKey As RegKey
  Set rKey = GetRegistryKeyA(hKey, Key)
  If Not rKey Is Nothing Then
    If Not rKey.Values.Item(ValueName) Is Nothing Then
      rKey.Values.Remove ValueName
    End If
  End If
  DeleteRegistryValueA = (GetRegistryValueObjectA(hKey, Key, ValueName) Is Nothing)
End Function

Public Function DeleteRegistryKeyA(hKey As REGISTRY_HKEYS, Key As String) As Boolean
  Dim rKey As RegKey
  Set rKey = GetRegistryKeyA(hKey, Key)
  If Not rKey Is Nothing Then
    Set rKey = rKey.Parent
    rKey.SubKeys.Remove GetFileNameA(Key)
  End If
  DeleteRegistryKeyA = (GetRegistryKeyA(hKey, Key) Is Nothing)
End Function

Public Function GetRegistryValueObjectA(hKey As REGISTRY_HKEYS, Key As String, ValueName As String) As RegValue
  Dim rKey As RegKey
  Set rKey = GetRegistryKeyA(hKey, Key)
  If Not rKey Is Nothing Then
    Set GetRegistryValueObjectA = rKey.Values.Item(ValueName)
  End If
End Function
