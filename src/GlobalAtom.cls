VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalAtom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_AtomID As Long

Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Long
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Long) As Long

Friend Property Get Value() As Long
  Value = m_AtomID
End Property

Private Sub Class_Initialize()
  m_AtomID = GlobalAddAtom(CStr(Now))
End Sub

Private Sub Class_Terminate()
  GlobalDeleteAtom m_AtomID
End Sub
