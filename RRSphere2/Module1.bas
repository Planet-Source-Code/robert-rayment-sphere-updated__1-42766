Attribute VB_Name = "Module1"
'Module1.bas

Option Base 1
DefBool A
DefByte B
DefLng C-T
DefSng U-Z
' $ strings
'------------------------------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

'------------------------------------------------------------------------------

' -----------------------------------------------------------
' APIs for getting bitmap bits to Memory

Public Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal HDC As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
(ByVal HDC As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
(ByVal HDC As Long) As Long

'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 4 x Pixel width here
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, usually 24 or 32
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type
Public bmp As BITMAP

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   'Colors(0 To 255) As RGBQUAD
End Type
Public bm As BITMAPINFO

' For transferring drawing in memory to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal HDC As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

'wUsage is one of:-
Public Const DIB_PAL_COLORS = 1 '  uses system
Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD
'dwRop is vbSrcCopy
'------------------------------------------------------------------------------
Public BScanLine
Public picWorldmem() As Long
Public picMapmem() As Long   ' Gets map
Public picLine() As Long      ' saves a line

Public picWorldWd, picWorldHt
Public picMapWd, picMapHt

Public ADone As Boolean

' Set at PRECALC
Public xindent()
'Public xtheta()
Public dFactor()
Public ixsc, ixdc    ' Source & Dest x centres
Public zPropFactor   ' Proportionality factor
Public zcos          ' = Cos(Tilt * pi# / 180)
Public zsin          ' = Sin(Tilt * pi# / 180)

' Set at Form_Load
Public Speed, Tilt

Public Const pi# = 3.14159265

Public Sub FillBMPStruc(ByVal iwidth As Long, ByVal iheight As Long)
With bm.bmiH
  .biSize = 40&
  .biwidth = iwidth
  .biheight = iheight
  .biPlanes = 1
  .biBitCount = 32      ' always 32 in this prog
  .biCompression = 0&
  ScanLine = (((iwidth * .biBitCount) + 31) \ 32) * 4
  ' Ensure expansion to 4B boundary
  ScanLine = (Int((ScanLine + 3) \ 4)) * 4

  .biSizeImage = ScanLine * Abs(.biheight)
  .biXPelsPerMeter = 0&
  .biYPelsPerMeter = 0&
  .biClrUsed = 0&
  .biClrImportant = 0&
End With
End Sub

Public Sub GETDIBS(ByVal PICIM As Long)

' PICIM is picbox.Image - handle to picbox memory
' from which pixels will be extracted and
' stored in picMapmem()

On Error GoTo DIBError

'Get info on picture loaded into PIC
GetObjectAPI PICIM, Len(bmp), bmp

NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)

FillBMPStruc bmp.bmWidth, bmp.bmHeight

' Set PicMem to receive color bytes or indexes or bits
picMapHt = bmp.bmHeight
picMapWd = bmp.bmWidth

ReDim picMapmem(1 To picMapWd, 1 To picMapHt)

' Load color bytes to PicMapmem
res = GetDIBits(NewDC, PICIM, 1, picMapHt, picMapmem(1, 1), bm, 1)

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC

Exit Sub
'==========
DIBError:
  MsgBox "DIB Error in GETDIBS", , "Sphere"
  DoEvents
  Unload Form1
  End
End Sub

Public Function zATan2(ByVal zy As Single, ByVal zx As Single) As Single
' Find angle Atan from -pi#/2 to +pi#/2
' Public pi#
If zx <> 0 Then
   zATan2 = Atn(zy / zx)
   If (zx < 0) Then
      If (zy < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
   End If
Else  ' zx=0
   If Abs(zy) > Abs(zx) Then   'Must be an overflow
      If zy > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
   Else
      zATan2 = 0   'Must be an underflow
   End If
End If
End Function


