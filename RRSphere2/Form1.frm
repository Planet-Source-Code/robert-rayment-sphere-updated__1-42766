VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Planet"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   679
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLatitude 
      Caption         =   "LAT Shift"
      Height          =   360
      Left            =   9030
      TabIndex        =   10
      Top             =   1695
      Width           =   1020
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   195
      LargeChange     =   3
      Left            =   9165
      Max             =   45
      Min             =   -45
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   180
      LargeChange     =   2
      Left            =   9105
      Max             =   8
      Min             =   -8
      TabIndex        =   4
      Top             =   450
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GO"
      Height          =   360
      Left            =   9030
      TabIndex        =   3
      Top             =   1290
      Width           =   1020
   End
   Begin VB.CommandButton Command5 
      Caption         =   "STOP"
      Height          =   360
      Left            =   9030
      TabIndex        =   2
      Top             =   2100
      Width           =   1020
   End
   Begin VB.PictureBox picWorld 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   240
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   120
      Width           =   2700
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   2400
      Left            =   3375
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   105
      Width           =   4800
   End
   Begin VB.Label Label4 
      Caption         =   "Tilt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8475
      TabIndex        =   9
      Top             =   930
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Speed"
      Height          =   225
      Left            =   8460
      TabIndex        =   8
      Top             =   465
      Width           =   510
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   195
      Left            =   9135
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Left            =   9090
      TabIndex        =   6
      Top             =   195
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sphere wrap, tilt & rotate  by  Robert Rayment

' Update with motion attached to scroll bars

' Update with Latitude shift & Byte to Long Array

Option Base 1

DefBool A
DefByte B
DefLng C-T
DefSng U-Z
' $ strings

Dim AScroller

Private Sub Form_Load()

Form1.Caption = "Sphere wrap, tilt & rotate  by  Robert Rayment"
Show
DoEvents


If (App.LogMode <> 1) Then
    MsgBox "Faster if compiled", vbExclamation, "Sphere"
End If

picMap.Picture = LoadPicture("map.jpg")
picMap.Refresh

' Ensure picWorld a square
picWorld.Height = picMap.Height
picWorld.Width = picWorld.Height
' Set Wd & Ht variables
picWorldWd = picWorld.Width
picWorldHt = picWorld.Height
picMapWd = picMap.Width
picMapHt = picMap.Height

' Set initial speed & tilt
Speed = 1   ' 1 pixel shift per frame
Tilt = 0    ' 0 degrees
zcos = Cos(Tilt * pi# / 180)
zsin = Sin(Tilt * pi# / 180)

AScroller = False

HScroll1.Value = Speed
Label1.Caption = HScroll1.Value
HScroll2.Value = Tilt
Label2.Caption = HScroll2.Value

AScroller = True

' Fill picMapmem with Map
GETDIBS picMap.Image

' Set up output array
ReDim picWorldmem(picWorldWd, picWorldHt)
' with its own bitmap struc for StretchDIBits
FillBMPStruc picWorldWd, picWorldHt

' For saving a line with latitude shift
ReDim picLine(picMapWd)

PRECALC

End Sub

Private Sub Command3_Click()
' GO

ADone = False

   Do
      
      Move_World_TILT
      
      ixsc = ixsc + Speed
      If ixsc > picMapWd Then ixsc = 1
      If ixsc < 1 Then ixsc = picMapWd
      
      DoEvents
   
   Loop Until ADone


End Sub

Private Sub Move_World_DIB()

' For Tilt=0 But NOT USED
' d - dest,  s - source,  c - centre

For iy = 1 To picWorldHt
   For zxd = xindent(iy) To picWorldWd - xindent(iy)
      
      isub = Int(zxd + 1)
      If isub > picWorldWd Then isub = picWorldWd
      ixs = ixsc + dFactor(isub)
      
      ixd = zxd
      If ixd < 1 Then ixd = 1
      If ixd > picWorldWd Then ixd = picWorldWd
      
      If ixs < 1 Then ixs = picMapWd + ixs
      If ixs > picMapWd Then ixs = ixs - picMapWd

      picWorldmem(ixd, iy) = picMapmem(ixs, iy)
   
   Next zxd
Next iy

ShowPicture
End Sub

Private Sub Move_World_TILT()
' Tilt angle degrees

'zcos = Cos(Tilt * pi# / 180)
'zsin = Sin(Tilt * pi# / 180)

' d - dest,  s - source,  c - centre

For iy = 1 To picWorldHt
   For zxd = xindent(iy) To picWorldWd - xindent(iy)

      ' Find rotated coords
      ixss = ixdc + (zxd - ixdc) * zcos - (iy - picMapHt / 2) * zsin
      iyss = picMapHt / 2 + (iy - picMapHt / 2) * zcos + (zxd - ixdc) * zsin
      If iyss > picMapHt Then iyss = picMapHt
      
      isub = Int(zxd + 1)
      If isub > picWorldWd Then isub = picWorldWd
      ixss = ixsc + dFactor(isub)
      
      ixd = zxd
      If ixd < 1 Then ixd = 1
      
      If ixss < 1 Then ixss = picMapWd + ixss
      If ixss > picMapWd Then ixss = ixss - picMapWd
      
      picWorldmem(ixd, iy) = picMapmem(ixss, iyss)

   Next zxd
Next iy

ShowPicture
End Sub

Private Sub ShowPicture()

bm.bmiH.biwidth = picWorldWd

If StretchDIBits(picWorld.HDC, _
   0&, 0&, picWorldWd, picWorldHt, _
   0&, 0&, picWorldWd, picWorldHt, _
   picWorldmem(1, 1), bm, _
   DIB_PAL_COLORS, vbSrcCopy) = 0 Then
      ADone = True
      MsgBox "Blit Error", , "Sphere"
      DoEvents
      Erase picWorldmem(), picMapmem()
      Unload Me
      End
End If
picWorld.Refresh
End Sub

Private Sub cmdLatitude_Click()
'res = CopyMemory(dest, src, nob)
' save bottom line

For L = 1 To 4

   CopyMemory picLine(1), picMapmem(1, 1), 4 * picMapWd
   
   ' shift pic down
   
   CopyMemory picMapmem(1, 1), picMapmem(1, 2), 4 * picMapWd * (picMapHt - 1)
   
   ' restore top line from bottom line
   
   CopyMemory picMapmem(1, picMapHt), picLine(1), 4 * picMapWd

Next L

picMap.Cls

bm.bmiH.biwidth = picMapWd

If StretchDIBits(picMap.HDC, _
   0&, 0&, picMapWd, picMapHt, _
   0&, 0&, picMapWd, picMapHt, _
   picMapmem(1, 1), bm, _
   DIB_PAL_COLORS, vbSrcCopy) = 0 Then
      ADone = True
      MsgBox "Blit Error", , "Sphere"
      DoEvents
      Erase picWorldmem(), picMapmem()
      Unload Me
      End
End If

picMap.Refresh
End Sub
Private Sub Command5_Click()
' STOP
ADone = True
End Sub

Private Sub PRECALC()
ReDim xindent(1 To picWorldHt)
ReDim dFactor(1 To picWorldHt)

' Basic maths
' Take a horizontal disc thru sphere

' Width of map(picWorldWd) /pi# =
' delta x from centre of disk/theta

' where theta is the angle made by the disk's radius as
' x-steps across disk.

' dFactor = theta * (picWorldWd) / pi#

' Spheroid radius
zR = 0.5 * picWorldWd

' Source & dest x-centre coords
ixsc = picMapWd \ 2
ixdc = picWorldWd \ 2

' Pre-calc horizontal slice's
' indentation from edge of rectangle & dFactors

For iy = 0 To picWorldHt - 1
   xindent(iy + 1) = zR - Sqr(iy * (2 * zR - iy))
Next iy

zradsq = zR * zR
For ixd = 0 To picWorldWd - 1
   xx = ixdc - ixd
   zz = Sqr(zradsq - (xx * xx))
   dFactor(ixd + 1) = Int(zATan2(zz, xx) * picWorldWd / pi# + 0.5)
Next ixd

End Sub

Private Sub Form_Unload(Cancel As Integer)
ADone = True
DoEvents
Erase picWorldmem(), picMapmem()
Unload Me
End
End Sub

Private Sub HScroll1_Change()
If AScroller = False Then Exit Sub
Speed = HScroll1.Value
Label1.Caption = Speed
End Sub

Private Sub HScroll1_Scroll()
If AScroller = False Then Exit Sub
Speed = HScroll1.Value
Label1.Caption = Speed
End Sub

Private Sub HScroll2_Change()
If AScroller = False Then Exit Sub
Tilt = HScroll2.Value
zcos = Cos(Tilt * pi# / 180)
zsin = Sin(Tilt * pi# / 180)
Label2.Caption = Tilt
End Sub

Private Sub HScroll2_Scroll()
If AScroller = False Then Exit Sub
Tilt = HScroll2.Value
zcos = Cos(Tilt * pi# / 180)
zsin = Sin(Tilt * pi# / 180)
Label2.Caption = Tilt
End Sub
