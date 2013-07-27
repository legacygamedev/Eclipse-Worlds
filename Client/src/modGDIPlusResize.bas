Attribute VB_Name = "modGDIPlusResize"
'Credits to FireXtol from vbforums.com
Option Explicit

' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const UnitPixel = 2

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function TransparentBlt Lib "msimg32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' PRIVATE GDI Declarations
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Type Guid
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type

Private Type PICTDESC
   size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type


Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long



Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'PRIVATE GDI+ Declarations
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal image As Long, Height As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long

Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As Long) As Long
Public Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As Long, ByVal srcImage As Long, dstImage As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal pImage As Long, ByVal x As Long, ByVal y As Long) As Long

Private m_lGDIpSmoothMode As Long

Public Type tFXDCS 'you need to store these
  hdc As Long 'you need these
  hBitmap As Long 'you need these
End Type


' Initialises GDI Plus
Public Function InitGDIPlus() As Long
    Dim token    As Long
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup token, gdipInit, ByVal 0&
    InitGDIPlus = token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(token As Long)
    GdiplusShutdown token
End Sub

Public Function LoadPictureGDIPlus(ByVal PicFile As String, Optional ByVal AutoLoad As Boolean = True, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal BackColor As Long, Optional ByVal RetainRatio As Boolean = False, Optional ByVal UseAlpha As Boolean = False) As IPicture
    Dim hdc     As Long
    Dim hBitmap As Long
    Dim croppedBitmap As Long
    Dim Img     As Long
    Dim token   As Long
    Dim tmpx    As Long
    Dim tmpy    As Long
    Dim realWidth As Long
    Dim realHeight As Long
   On Error GoTo LoadPictureGDIPlus_Error

    ' Load the image
    If Len(Mid$(PicFile$, InStrRev(PicFile$, "\"))) < 2 Then Exit Function
    
    InitGDIPlus
    
    If AutoLoad Then token = InitGDIPlus
    hdc = GdipLoadImageFromFile(StrPtr(GetShortName(PicFile$)), Img)
    If hdc <> 0 Then
        GdipDisposeImage Img
        If AutoLoad Then FreeGDIPlus token
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    GdipGetImageWidth Img, tmpx
    GdipGetImageHeight Img, tmpy

    If tmpx < Width Then Width = tmpx
    If tmpy < Height Then Height = tmpy
    
     GdipGetImageWidth Img, realWidth
     GdipGetImageHeight Img, realHeight
    
    If Width = -1 Then
        Width = realWidth
    End If
    If Height = -1 Then
        Height = realHeight
    End If
    
    ' Initialise the hDC
    InitDC hdc, hBitmap, BackColor, Width, Height, UseAlpha

    ' Resize the picture
    gdipResize Img, hdc, Width, Height, RetainRatio
    If realWidth > Width Or realHeight > Height Then
        GdipCloneBitmapAreaI -32, 0, 32, 32, 0, Img, croppedBitmap
        Img = croppedBitmap
    End If
    GdipDisposeImage Img
    
    'TransparentBlt Form1.hdc, 0, 0, Width, Height, hdc, 0, 0, Width, Height, 0
    
    'AlphaBlt Form1.hdc, 0, 0, Width, Height, hdc, 0, 0, Width, Height, , True
    
    'select bitmap out of DC
    hBitmap = SelectObject(hdc, hBitmap)
    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
    If AutoLoad Then FreeGDIPlus token
    DeleteDC hdc
    
    'should probably do this eventually: DeleteObject hBitmap

   On Error GoTo 0
   Exit Function
   
    
LoadPictureGDIPlus_Error:
    GdipDisposeImage Img
    DeleteObject SelectObject(hdc, hBitmap)
    DeleteDC hdc

End Function

' Loads the picture (optionally resized)
Public Sub LoadPictureFXDC(ByVal PicFile As String, fxDC As tFXDCS, Optional ByVal AutoLoad As Boolean = True, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal BackColor As Long, Optional ByVal RetainRatio As Boolean = False, Optional ByVal UseAlpha As Boolean = False)
'purpose: load graphic from file
'returns: fxDC, passed by reference, return a hDC and a stock bitmap handle to *reselect into DC to return bitmap this creates*
'function returns: .Picture compatible object(like form, picturebox, etc)
    Dim hdc     As Long
    Dim hBitmap As Long
    Dim Img     As Long
    Dim token   As Long
    Dim tmpx    As Long
    Dim tmpy    As Long
   On Error GoTo LoadPictureFXDC_Error

    ' Load the image
    If Len(Mid$(PicFile$, InStrRev(PicFile$, "\"))) < 2 Then Exit Sub

    If AutoLoad Then token = InitGDIPlus
    If GdipLoadImageFromFile(StrPtr(GetShortName(PicFile$)), Img) <> 0 Then
        GdipDisposeImage Img
        If AutoLoad Then FreeGDIPlus token
        Exit Sub
    End If
    
    ' Calculate picture's width and height if not specified
    GdipGetImageWidth Img, tmpx
    GdipGetImageHeight Img, tmpy

    If tmpx < Width Then Width = tmpx
    If tmpy < Height Then Height = tmpy
    
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
    
    ' Initialise the hDC
    InitDC hdc, hBitmap, BackColor, Width, Height, UseAlpha

    ' Resize the picture
    gdipResize Img, hdc, Width, Height, RetainRatio
    GdipDisposeImage Img
    fxDC.hdc = hdc
    fxDC.hBitmap = hBitmap
    
    'select bitmap out of DC
    'hBitmap = SelectObject(hDC, hBitmap)

    ' Create the picture
    'Set LoadPictureGDIPlus = CreatePicture(hBitmap)
    If AutoLoad Then FreeGDIPlus token
    'DeleteDC hDC
    
    'should probably do this eventually: DeleteObject hBitmap

   On Error GoTo 0
   Exit Sub
   
    
LoadPictureFXDC_Error:
    GdipDisposeImage Img
    DeleteObject SelectObject(hdc, hBitmap)
    DeleteDC hdc

End Sub

' Initialises the hDC to draw
Public Sub InitDC(hdc As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long, ByVal UseAlpha As Boolean)
    Dim hBrush As Long

    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hdc = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hdc, PLANES), GetDeviceCaps(hdc, BITSPIXEL), ByVal 0&)
    If hBitmap <> 0 Then
      hBitmap = SelectObject(hdc, hBitmap)
      If Not UseAlpha Then
          hBrush = CreateSolidBrush(BackColor)
          hBrush = SelectObject(hdc, hBrush)
          PatBlt hdc, 0, 0, Width, Height, PATCOPY
          DeleteObject SelectObject(hdc, hBrush)
      End If
    Else
      DeleteDC hdc
      Err.Raise 1, "InitDC", "Bitmap creation failed"
    End If
End Sub

Public Sub UnloadFXDC(fxDC As tFXDCS)
DeleteObject SelectObject(fxDC.hdc, fxDC.hBitmap)
DeleteDC fxDC.hdc
fxDC.hdc = 0
fxDC.hBitmap = 0
End Sub

' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hdc As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hdc, graphics
    GdipSetInterpolationMode graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, 32, 32, UnitPixel, 0, 0, 0
    Else

        GdipDrawImageI graphics, Img, -32, 0
    End If
    GdipDeleteGraphics graphics
End Sub

' Creates a VB compatible Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
    Dim IID_IDispatch As Guid
    Dim Pic           As PICTDESC
    Dim IPic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
        
    ' Fill Pic with necessary parts
    Pic.size = Len(Pic)        ' Length of structure
    Pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    Pic.hBmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
       
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    
    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)

End Function

Public Sub AlphaBlt(ByVal dhDC As Long, ByVal dx As Long, ByVal dy As Long, ByVal dW As Long, ByVal dH As Long, ByVal shDC As Long, ByVal Sx As Long, ByVal Sy As Long, ByVal sW As Long, ByVal sH As Long, Optional ByVal UseAlpha As Boolean, Optional ByVal AlphaConstant As Byte = 255)
Dim tmpHDC As Long, hBitmap As Long, bitmap As Long
Dim Blend As BLENDFUNCTION, BlendLng As Long

If UseAlpha Then Blend.AlphaFormat = 1 'use source alpha
Blend.SourceConstantAlpha = AlphaConstant
CopyMemory BlendLng, Blend, 4
    
AlphaBlend dhDC, dx, dy, dW, dH, shDC, Sx, Sy, dW, dH, BlendLng
End Sub

Public Sub AlphaTransBlt(ByVal dhDC As Long, ByVal dx As Long, ByVal dy As Long, ByVal dW As Long, ByVal dH As Long, ByVal shDC As Long, ByVal Sx As Long, ByVal Sy As Long, ByVal sW As Long, ByVal sH As Long, Optional ByVal TransColor As Long = 65024, Optional ByVal AlphaConstant As Byte = 255, Optional ByVal PerPixelAlpha As Boolean = False)
Dim tmpHDC As Long, hBitmap As Long, bitmap As Long
Dim Blend As BLENDFUNCTION, BlendLng As Long
    

Blend.SourceConstantAlpha = AlphaConstant
If PerPixelAlpha Then Blend.AlphaFormat = 1
CopyMemory BlendLng, Blend, 4
    
If AlphaConstant = 255 Then
    If PerPixelAlpha Then
    AlphaBlend dhDC, dx, dy, dW, dH, shDC, Sx, Sy, dW, dH, BlendLng
    Else
    TransparentBlt dhDC, dx, dy, dW, dH, shDC, Sx, Sy, sW, sH, TransColor
    End If
Else
    InitDC tmpHDC, hBitmap, TransColor, dW, dH, True
    BitBlt tmpHDC, 0, 0, dW, dH, dhDC, dx, dy, vbSrcCopy 'blt the background on the destination DC to a temporary DC
    If PerPixelAlpha Then
        Blend.SourceConstantAlpha = 255
        CopyMemory BlendLng, Blend, 4
        AlphaBlend tmpHDC, 0, 0, dW, dH, shDC, Sx, Sy, dW, dH, BlendLng
    Else
        TransparentBlt tmpHDC, 0, 0, dW, dH, shDC, Sx, Sy, dW, dH, TransColor
    End If

    Blend.SourceConstantAlpha = AlphaConstant
    Blend.AlphaFormat = 0 'disable source alpha
    CopyMemory BlendLng, Blend, 4
    
    AlphaBlend dhDC, dx, dy, dW, dH, tmpHDC, 0, 0, dW, dH, BlendLng 'handles alphaconstant
    DeleteObject SelectObject(tmpHDC, hBitmap)
    DeleteDC tmpHDC
End If

End Sub

Public Property Get GDIpSmoothMode() As Long

  GDIpSmoothMode = m_lGDIpSmoothMode

End Property

Public Property Let GDIpSmoothMode(ByVal lGDIpSmoothMode As Long)

  m_lGDIpSmoothMode = lGDIpSmoothMode

End Property
