VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum REGISTRY_HKEYS
  HKEY_CLASSES_ROOT = RegObj.HKEY_CLASSES_ROOT
  HKEY_CURRENT_USER = RegObj.HKEY_CURRENT_USER
  HKEY_LOCAL_MACHINE = RegObj.HKEY_LOCAL_MACHINE
  HKEY_USERS = RegObj.HKEY_USERS
  HKEY_PERFORMANCE_DATA = RegObj.HKEY_PERFORMANCE_DATA
End Enum

Public Function GetRegistryKey(hKey As REGISTRY_HKEYS, Optional Key As String, Optional ForceCreate As Boolean) As RegKey
  Set GetRegistryKey = GetRegistryKeyA(hKey, Key, ForceCreate)
End Function

Public Function CreateRegistryKey(hKey As REGISTRY_HKEYS, Key As String) As RegKey
  Set CreateRegistryKey = CreateRegistryKeyA(hKey, Key)
End Function

Public Function GetRegistryValue(hKey As REGISTRY_HKEYS, Key As String, Optional ValueName As String) As Variant
  GetRegistryValue = GetRegistryValueA(hKey, Key, ValueName)
End Function

Public Function SetRegistryValue(NewValue As Variant, hKey As REGISTRY_HKEYS, Key As String, Optional ValueName As String, Optional ForceCreate As Boolean) As Boolean
  SetRegistryValue = SetRegistryValueA(NewValue, hKey, Key, ValueName, ForceCreate)
End Function

Public Function DeleteRegistryValue(hKey As REGISTRY_HKEYS, Key As String, ValueName As String) As Boolean
  DeleteRegistryValue = DeleteRegistryValueA(hKey, Key, ValueName)
End Function

Public Function DeleteRegistryKey(hKey As REGISTRY_HKEYS, Key As String) As Boolean
  DeleteRegistryKey = DeleteRegistryKeyA(hKey, Key)
End Function

Public Function GetRegistryValueObject(hKey As REGISTRY_HKEYS, Key As String, ValueName As String) As RegValue
  Set GetRegistryValueObject = GetRegistryValueObjectA(hKey, Key, ValueName)
End Function
