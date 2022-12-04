VERSION 5.00
Begin VB.Form aerofrm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   744
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tcount 
      Interval        =   1
      Left            =   0
      Top             =   720
   End
   Begin VB.PictureBox pngcon 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11895
      TabIndex        =   0
      Top             =   120
      Width           =   11895
   End
End
Attribute VB_Name = "aerofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Karmba_a@hotmail.com
' MSLE 2007
Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    CX As Long
    CY As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Dim blendFunc32bpp As BLENDFUNCTION
Dim mDC As Long
Dim mainBitmap As Long
Dim oldBitmap As Long
Dim token As Long

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Dim FileRes As Integer
Dim Buffer() As Byte
Function Function_End()
    Call GdiplusShutdown(token)
    SelectObject mDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    Unload frmmain
    Unload Me
End Function
Function Cool_Memory()
    SelectObject mDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
End Function
Private Function MakeTrans(pngPath As String) As Boolean
   Dim tempBI As BITMAPINFO
   Dim tempBlend As BLENDFUNCTION
   Dim lngHeight As Long, lngWidth As Long
   Dim curWinLong As Long
   Dim img As Long
   Dim graphics As Long
   Dim winSize As Size
   Dim srcPoint As POINTAPI
   
   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32
      .biHeight = Me.ScaleHeight
      .biWidth = Me.ScaleWidth
      .biPlanes = 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
   End With
   mDC = CreateCompatibleDC(Me.hdc)
   mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
   oldBitmap = SelectObject(mDC, mainBitmap)

   Call GdipCreateFromHDC(mDC, graphics)
   Call GdipLoadImageFromFile(StrConv(pngPath, vbUnicode), img)
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)

   curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
   SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
   
   srcPoint.X = 0
   srcPoint.Y = 0
   winSize.CX = Me.ScaleWidth
   winSize.CY = Me.ScaleHeight
    
   With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = 255
   End With
    
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
   Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Function
Sub MoveForm(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub
Sub Center(FormName As Form)
Move (Screen.Width - FormName.Width) \ 2, (Screen.Height - FormName.Height) \ 2
End Sub
Private Sub Form_Initialize()
   Dim GpInput As GdiplusStartupInput
   Dim guiloc As String
   guiloc = App.Path & "\gui\gui_main.png"
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) <> 0 Then
     MsgBox "Error loading GDI+!", vbCritical
     Function_End
   End If
   MakeTrans (guiloc)
   Cool_Memory
End Sub

Private Sub Form_Load()
    Me.Height = 7680
    Me.Width = 11550
Call Center(aerofrm)
Cool_Memory
End Sub
Private Sub Form_Unload(Cancel As Integer)
Function_End
End Sub

Private Sub pngcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm Me
    Cool_Memory
End Sub

Private Sub tcount_Timer()
On Error Resume Next
    Cool_Memory
    frmmain.Move (aerofrm.Left) + 198, (aerofrm.Top) + 220
    frmmain.Show
End Sub
