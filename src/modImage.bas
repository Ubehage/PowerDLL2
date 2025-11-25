Attribute VB_Name = "modImage"
Option Explicit

Public Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Public Type BITMAPFILEHEADER
  bfType As Integer
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type

Private Declare Function InvertGraphicsRect Lib "user32" Alias "InvertRect" (ByVal hDC As Long, lpRect As RECT) As Long

Public Function ResizePictureA(Picture As IPictureDisp, Width As Long, Height As Long) As IPictureDisp
  If Not Picture Is Nothing Then
    Load frmObjects
    With frmObjects.picPicture
      .Width = ((.Width - .ScaleWidth) + (Screen.TwipsPerPixelX * Width))
      .Height = ((.Height - .ScaleHeight) + (Screen.TwipsPerPixelY * Height))
      .PaintPicture Picture, 0, 0, .ScaleWidth, .ScaleHeight
      .Picture = .Image
      Set ResizePictureA = .Picture
    End With
    Unload frmObjects
    Set frmObjects = Nothing
  End If
End Function

Public Function GetBitmapPictureSizeA(BitmapFile As String) As IMAGE_SIZE
  On Error GoTo Error1
  Dim fHeader As BITMAPFILEHEADER
  Dim iHeader As BITMAPINFOHEADER
  Dim fFile As Integer
  fFile = FreeFile
  Open BitmapFile For Binary Access Read As fFile
  Get #fFile, , fHeader
  Get #fFile, , iHeader
  With GetBitmapPictureSizeA
    .Width = iHeader.biWidth
    .Height = iHeader.biHeight
  End With
ExitNow:
  Close #fFile
  On Error GoTo 0
  Exit Function
Error1:
  Resume ExitNow
End Function

Public Function GetJPEGPictureSizeA(JPEGFile As String) As IMAGE_SIZE
  On Error GoTo Error1
  Dim bChar As Byte
  Dim a As Byte
  Dim b As Byte
  Dim c As Byte
  Dim d As Byte
  Dim e As Byte
  Dim f As Byte
  Dim i As Integer
  Dim DotPos As Integer
  Dim Header As String
  Dim blExit As Boolean
  Dim MarkerLen As Long
  Dim marker As String
  Dim imgWidth As Integer
  Dim imgHeight As Integer
  Dim imgSize As String
  Dim fFile As Integer
  fFile = FreeFile
  Open JPEGFile For Binary Access Read As fFile
  imgSize = (LOF(fFile) / 1024)
  DotPos = InStr(imgSize, ",")
  imgSize = Left(imgSize, DotPos)
  Get #fFile, , bChar
  Header = Hex(bChar)
  Get #fFile, , bChar
  Header = Header & Hex(bChar)
  If Not Header = "FFD8" Then
    GoTo ExitNow
  End If
  While Not blExit
    Do Until Hex(bChar) = "FF"
      Get #fFile, , bChar
    Loop
    Get #fFile, , bChar
    If (Hex(bChar) >= "C0" And Hex(bChar) <= "C3") Then
      Get #fFile, , bChar
      Get #fFile, , bChar
      Get #fFile, , bChar
      Get #fFile, , bChar
      a = bChar
      Get #fFile, , bChar
      b = bChar
      Get #fFile, , bChar
      c = bChar
      Get #fFile, , bChar
      d = bChar
      imgHeight = CInt(((a * 256) + b))
      imgWidth = CInt(((c * 256) + d))
      blExit = True
    Else
      If Hex(bChar) = "DA" Then
        blExit = True
      Else
        Get #fFile, , bChar
        e = bChar
        Get #fFile, , bChar
        f = bChar
        MarkerLen = (((e * 256) + f) - 2)
        marker = String(MarkerLen, vbNullChar)
        Get #fFile, , marker
      End If
    End If
  Wend
  With GetJPEGPictureSizeA
    .Width = imgWidth
    .Height = imgHeight
  End With
ExitNow:
  Close #1
  On Error GoTo 0
  Exit Function
Error1:
  Resume ExitNow
End Function

Public Function GetGIFPictureSizeA(GIFFile As String) As IMAGE_SIZE
  On Error GoTo Error1
  Dim bChar As Byte
  Dim i As Integer
  Dim DotPos As Integer
  Dim Header As String
  Dim blExit As Boolean
  Dim a As String
  Dim b As String
  Dim imgWidth As Integer
  Dim imgHeight As Integer
  Dim imgSize As String
  Dim fFile As Integer
  fFile = FreeFile
  Open GIFFile For Binary Access Read As fFile
  imgSize = (LOF(fFile) / 1024)
  DotPos = InStr(imgSize, ",")
  imgSize = Left(imgSize, (DotPos - 1))
  For i = 0 To 5
    Get #fFile, , bChar
    Header = (Header + Chr(bChar))
  Next
  If Not Left(Header, 3) = "GIF" Then
    GoTo ExitNow
  End If
  Get #fFile, , bChar
  a = (a + Chr(bChar))
  Get #fFile, , bChar
  a = (a + Chr(bChar))
  imgWidth = CInt(Asc(Left(a, 1)) + 256 * Asc(Right(a, 1)))
  Get #fFile, , bChar
  b = (b + Chr(bChar))
  Get #fFile, , bChar
  b = (b + Chr(bChar))
  imgHeight = CInt(Asc(Left(b, 1)) + 256 * Asc(Right(b, 1)))
  With GetGIFPictureSizeA
    .Width = imgWidth
    .Height = imgHeight
  End With
ExitNow:
  Close #1
  On Error GoTo 0
  Exit Function
Error1:
  Resume ExitNow
End Function

Public Function GetPictureSizeFromFileA(PictureFile As String) As IMAGE_SIZE
  Select Case LCase(GetFileExtensionA(PictureFile))
    Case "bmp"
      GetPictureSizeFromFileA = GetBitmapPictureSizeA(PictureFile)
    Case "gif"
      GetPictureSizeFromFileA = GetGIFPictureSizeA(PictureFile)
    Case "jpeg", "jpg"
      GetPictureSizeFromFileA = GetJPEGPictureSizeA(PictureFile)
    Case "ico", "cur", "ani"
      With GetPictureSizeFromFileA
        .Width = 32
        .Height = 32
      End With
  End Select
End Function

Public Function IsPictureFileA(File As String) As Boolean
  With GetPictureSizeFromFileA(File)
    IsPictureFileA = Not (.Width = 0 Or .Height = 0)
  End With
End Function

Public Function PrintScreenA() As IPictureDisp
  Dim hWndDesk As Long
  Dim hDcDesk As Long
  Dim LeftDesk As Long
  Dim TopDesk As Long
  Dim WidthDesk As Long
  Dim HeightDesk As Long
  WidthDesk = (Screen.Width / Screen.TwipsPerPixelX)
  HeightDesk = (Screen.Height / Screen.TwipsPerPixelY)
  hWndDesk = GetDesktopWindow
  hDcDesk = GetWindowDC(hWndDesk)
  Load frmObjects
  With frmObjects.picPicture
    .AutoRedraw = True
    .Move 0, 0, ((.Width - .ScaleWidth) + Screen.Width), ((.Height - .ScaleHeight) + Screen.Height)
    BitBlt .hDC, 0, 0, WidthDesk, HeightDesk, hDcDesk, LeftDesk, TopDesk, vbSrcCopy
    .Picture = .Image
    Set PrintScreenA = .Picture
  End With
  Unload frmObjects
  Set frmObjects = Nothing
  ReleaseDC hWndDesk, hDcDesk
End Function

Public Function InvertColorA(Color As Long) As Long
  Dim Red As Integer
  Dim Green As Integer
  Dim Blue As Integer
  SplitRGBA Color, Red, Green, Blue
  Red = InvertColorValue(Red)
  Green = InvertColorValue(Green)
  Blue = InvertColorValue(Blue)
  InvertColorA = RGB(Red, Green, Blue)
End Function

Public Function AddColorsA(Color1 As Long, Color2 As Long) As Long
  Dim Red1 As Integer
  Dim Green1 As Integer
  Dim Blue1 As Integer
  Dim Red2 As Integer
  Dim Green2 As Integer
  Dim Blue2 As Integer
  SplitRGBA Color1, Red1, Green1, Blue1
  SplitRGBA Color2, Red2, Green2, Blue2
  Red1 = AddColorValues(Red1, Red2)
  Green1 = AddColorValues(Green1, Green2)
  Blue1 = AddColorValues(Blue1, Blue2)
  AddColorsA = RGB(Red1, Green1, Blue1)
End Function

Public Function InvertPictureA(Picture As Object) As Object
  Dim pRect As RECT
  Load frmObjects
  With frmObjects.picPicture
    .AutoSize = True
    .AutoRedraw = True
    .Picture = Picture
    pRect.Left = 0
    pRect.Top = 0
    pRect.Right = (.ScaleWidth / Screen.TwipsPerPixelX)
    pRect.Bottom = (.ScaleHeight / Screen.TwipsPerPixelY)
    InvertGraphicsRect .hDC, pRect
    .Picture = .Image
    Set InvertPictureA = .Picture
  End With
  Unload frmObjects
  Set frmObjects = Nothing
End Function
