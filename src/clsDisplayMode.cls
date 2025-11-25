VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisplayMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Bits As Long
Dim m_Width As Long
Dim m_Height As Long

Dim m_ModesParent As PowerDLL2.clsDisplayModes

Public Property Get Bits() As Long
  Bits = m_Bits
End Property
Public Property Get Width() As Long
  Width = m_Width
End Property
Public Property Get Height() As Long
  Height = m_Height
End Property
Public Property Get Colors() As Double
  Select Case m_Bits
    Case 4
      Colors = 16
    Case 8
      Colors = 256
    Case 16
      Colors = 65536
    Case 24
      Colors = 16777216
    Case 32
      Colors = 4294967296#
  End Select
End Property
Friend Property Get ModesParent() As PowerDLL2.clsDisplayModes
  Set ModesParent = m_ModesParent
End Property
Friend Property Let Bits(New_Bits As Long)
  m_Bits = New_Bits
End Property
Friend Property Let Width(New_Width As Long)
  m_Width = New_Width
End Property
Friend Property Let Height(New_Height As Long)
  m_Height = New_Height
End Property
Friend Property Let ModesParent(New_ModesParent As PowerDLL2.clsDisplayModes)
  Set m_ModesParent = New_ModesParent
End Property

Public Sub UseThisMode()
  ChangeDisplayModeA m_Width, m_Height, m_Bits
End Sub
