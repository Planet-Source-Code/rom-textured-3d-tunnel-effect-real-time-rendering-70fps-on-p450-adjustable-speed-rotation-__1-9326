VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Screensaver"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "vbRipple.frx":0000
   ScaleHeight     =   3375
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   2640
   End
   Begin VB.PictureBox say 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   600
      Picture         =   "vbRipple.frx":FE42
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos& Lib "user32.dll" (lpPoint As POINTAPI)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type pt
    X As Long
    Y As Long
End Type
Private Type POINTAPI
    X As Integer
    Z As Integer
    Y As Integer
End Type

Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const SRCCOPY = &HCC0020
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6

Public DD As DirectDraw
Private Lookup3(256) As Long
Private Lookup4(300, 125) As Long
Private Lookup5(300, 125) As Long
Private Lookup7(250) As pt
Private LSin(1000)
Private LCos(1000)

Private LookupRollIn(320, 200) As pt

Public wind, UpSpd%, AtNow%, Smooth As Boolean, fps%, SpinSpd%, AtNowSpin%
Private pt As POINTAPI

Sub DoEffect()
On Error Resume Next

Dim pict() As Byte
Dim pict2() As Byte

Dim sa As SAFEARRAY2D, bmp As BITMAP
Dim sa2 As SAFEARRAY2D, bmp2 As BITMAP

'info on bitmaps in each buffer
GetObjectAPI Form1.Picture, Len(bmp), bmp
GetObjectAPI say.Picture, Len(bmp2), bmp2

'must be 8bpp
If bmp.bmPlanes <> 1 Or bmp.bmBitsPixel <> 8 Then
    If wind = 0 Then
        ShowCursor 1
    End If
    End
End If

'point to pixels of each buffer
With sa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp.bmWidthBytes
    .pvData = bmp.bmBits
End With
CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4

With sa2
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp2.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp2.bmWidthBytes
    .pvData = bmp2.bmBits
End With
CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4

AtNow% = AtNow% + UpSpd%
If AtNow% > 254 Then
    AtNow% = AtNow% - 254
ElseIf AtNow% < 0 Then
    AtNow% = AtNow% + 254
End If

AtNowSpin% = AtNowSpin% + SpinSpd%
If AtNowSpin% > 254 Then
    AtNowSpin% = AtNowSpin% - 254
ElseIf AtNowSpin% < 0 Then
    AtNowSpin% = AtNowSpin% + 254
End If

For i% = 0 To 320 'UBound(pict, 1)
For X% = 0 To 200 'UBound(pict, 2)
    i2% = LookupRollIn(i%, X%).Y + AtNowSpin%
    If i2% > 254 Then
        i2% = i2% - 254
    ElseIf i2% < 0 Then
        i2% = i2% + 254
    End If
    pict(i%, X%) = pict2(LookupRollIn(i%, X%).X + AtNow%, i2%)
Next X%
Next i%

CopyMemory ByVal VarPtrArray(pict), 0&, 4
CopyMemory ByVal VarPtrArray(pict2), 0&, 4

fps% = fps% + 1
Form1.Refresh
End Sub


Private Sub Form_DblClick()
'end program
If wind = 0 Then
    ShowCursor 1
End If
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
    If UpSpd% >= 256 Then
        Exit Sub
    End If
    UpSpd% = UpSpd% + 1
ElseIf KeyCode = vbKeyDown Then
    If UpSpd% <= -256 Then
        Exit Sub
    End If
    UpSpd% = UpSpd% - 1
End If

If KeyCode = vbKeyLeft Then
    If SpinSpd% >= 256 Then
        Exit Sub
    End If
    SpinSpd% = SpinSpd% + 1
ElseIf KeyCode = vbKeyRight Then
    If SpinSpd% <= -256 Then
        Exit Sub
    End If
    SpinSpd% = SpinSpd% - 1
End If
End Sub

Private Sub Form_Load()

wind = 0 'determines whether program will run in window or fullscreen
UpSpd% = 3
SpinSpd% = 1
Smooth = False

Randomize

w = 320  '}___ width and height
H = 200  '}    of screen

If wind = 0 Then
    ShowCursor 0
End If

'Form1.Picture = LoadPicture(App.Path & "\bg.bmp") '(must be 320x200)
'say.Picture = LoadPicture(App.Path & "\plasma.bmp") '(must be 256x256)

Form1.Width = (w * Screen.TwipsPerPixelX)
Form1.Height = (H * Screen.TwipsPerPixelY)
say.Width = (w * Screen.TwipsPerPixelX)
say.Height = (H * Screen.TwipsPerPixelY)

If wind = 0 Then
    'direct draw stuff to change res
    DirectDrawCreate ByVal 0&, DD, Nothing
    DD.SetCooperativeLevel Me.hwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
    DD.SetDisplayMode 320, 200, 16
End If

Form1.Visible = True

If wind = 1 Then
    StayOnTop Me
End If

'Print "     Loading..."

SetLooks 'create lookup tables

'main loop
Plasma 'this will generate a texture and smooth it
Do
DoEvents
DoEffect
Loop
End Sub

Sub SetLooks()
'transformation lookup tables
For i = 0 To 320
For X = 0 To 200
    a = i - 160
    b = X - 100
    
    If b = 0 Then
        ANGLE = 90 * Sgn(a)
        r = 5000 / (Abs(a) + 0.0001)
    Else
        r = 5000 / (Sqr(Abs(a) ^ 2 + Abs(b) ^ 2))
        ANGLE = Atn(a / b) / (3.141592 / 180)
    End If
    
    If b < 0 Then
        ANGLE = ANGLE - 180
    End If
    
    ANGLE = (ANGLE + 360) * ((256) / 360)
    
    
    LookupRollIn(i, X).X = ((r) Mod (256))
    
    If LookupRollIn(i, X).X >= 256 Then
        LookupRollIn(i, X).X = 256
    End If
    If LookupRollIn(i, X).X < 0 Then
        LookupRollIn(i, X).X = 0
    End If
    
    LookupRollIn(i, X).Y = ((ANGLE) Mod (256 - 1)) + 1
    
    If LookupRollIn(i, X).Y >= 256 Then
        LookupRollIn(i, X).Y = 256
    End If
    If LookupRollIn(i, X).Y < 0 Then
        LookupRollIn(i, X).Y = 0
    End If
Next X
Next i
End Sub

Sub StayOnTop(TheForm As Form) 'make form stay on top, if window
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub Timer1_Timer()
Caption = fps%
fps% = 0
End Sub

Sub Plasma()
On Error Resume Next

Dim pict() As Byte

Dim sa As SAFEARRAY2D, bmp As BITMAP

'info on bitmaps in each buffer
GetObjectAPI say.Picture, Len(bmp), bmp

'must be 8bpp
If bmp.bmPlanes <> 1 Or bmp.bmBitsPixel <> 8 Then
    If wind = 0 Then
        ShowCursor 1
    End If
    End
End If

'point to pixels of each buffer
With sa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp.bmWidthBytes
    .pvData = bmp.bmBits
End With
CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4

'generate texture here

'blur (smooth)
For c% = 3 To UBound(pict, 1) - 3 Step 1
    For r% = 1 To UBound(pict, 2) - 1 Step 1
        pict(c%, r%) = (CInt(pict(c%, r%)) + _
        CInt(pict(c% - 1, r%)) + _
        CInt(pict(c% + 1, r%)) + _
        CInt(pict(c%, r% - 1)) + _
        CInt(pict(c%, r% + 1))) \ 5
    Next
Next

CopyMemory ByVal VarPtrArray(pict), 0&, 4
say.Refresh
End Sub
