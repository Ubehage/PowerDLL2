VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImage"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type IMAGE_SIZE
  Width As Long
  Height As Long
End Type

Public Function ResizePicture(Picture As IPictureDisp, Width As Long, Height As Long) As IPictureDisp
  Set ResizePicture = ResizePictureA(Picture, Width, Height)
End Function

Public Function GetBitmapPictureSize(BitmapFile As String) As IMAGE_SIZE
  GetBitmapPictureSize = GetBitmapPictureSizeA(BitmapFile)
End Function

Public Function GetJPEGPictureSize(JPEGFile As String) As IMAGE_SIZE
  GetJPEGPictureSize = GetJPEGPictureSizeA(JPEGFile)
End Function

Public Function GetGIFPictureSize(GIFFile As String) As IMAGE_SIZE
  GetGIFPictureSize = GetGIFPictureSizeA(GIFFile)
End Function

Public Function GetPictureSizeFromFile(PictureFile As String) As IMAGE_SIZE
  GetPictureSizeFromFile = GetPictureSizeFromFileA(PictureFile)
End Function

Public Function IsPictureFile(File As String) As Boolean
  IsPictureFile = IsPictureFileA(File)
End Function

Public Function PrintScreen() As Object
  Set PrintScreen = PrintScreenA
End Function

Public Function InvertColor(Color As Long) As Long
  InvertColor = InvertColorA(Color)
End Function

Public Function AddColors(Color1 As Long, Color2 As Long) As Long
  AddColors = AddColorsA(Color1, Color2)
End Function

Public Function InvertPicture(Picture As Object) As Object
  Set InvertPicture = InvertPictureA(Picture)
End Function
