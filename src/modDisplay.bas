Attribute VB_Name = "modDisplay"
Option Explicit

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Public Function ChangeDisplayModeA(Width As Long, Height As Long, Bits As Long) As Boolean
  Dim dMode As DEVMODE
  With dMode
    .dmPelsWidth = Width
    .dmPelsHeight = Height
    .dmBitsPerPel = Bits
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    .dmSize = LenB(dMode)
  End With
  ChangeDisplayModeA = (ChangeDisplaySettings(dMode, CDS_FORCE) = 0)
End Function
