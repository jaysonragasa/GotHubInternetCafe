Attribute VB_Name = "modShadow"
Option Explicit

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Const CLR_INVALID                    As Integer = -1

'Private Type RECT
'    Left                                     As Long
'    Top                                      As Long
'    Right                                    As Long
'    Bottom                                   As Long
'End Type

Public Sub DrawShadow(ByVal wForm As Form, ByVal ctrlParent As Object, ByVal ctrlDestDrawShadow As PictureBox, Optional ByVal Opacity As Integer = 4)
     Dim temp       As StdPicture
     
     ctrlDestDrawShadow.AutoRedraw = True
     ctrlDestDrawShadow.Move ctrlParent.Left + (4 * Screen.TwipsPerPixelX), ctrlParent.Top + (4 * Screen.TwipsPerPixelY), ctrlParent.Width, ctrlParent.Height
     ctrlDestDrawShadow.Visible = False
     DoEvents
     
     Set temp = CaptureForm(wForm)
     DoEvents
     
     ctrlDestDrawShadow.Picture = LoadPicture("")
     ctrlDestDrawShadow.PaintPicture temp, _
                                     -(ctrlDestDrawShadow.Left + (wForm.Width - wForm.ScaleWidth) - 0), _
                                     -(ctrlDestDrawShadow.Top + (wForm.Height - wForm.ScaleHeight) - 0)
     DoEvents
     
     ctrlDestDrawShadow.Visible = True
     
     Set temp = Nothing
     
     Call DrawMenuShadow(ctrlDestDrawShadow.hwnd, ctrlDestDrawShadow.hdc, Opacity)
End Sub

Public Sub DrawShadow_IC(ByVal wForm As Form, ByVal ctrlFrame As Object, ByVal ctrlParent As Object, ByVal ctrlDestDrawShadow As PictureBox, Optional ByVal Opacity As Integer = 4)
     Dim temp       As StdPicture
     
     ctrlDestDrawShadow.AutoRedraw = True
     ctrlDestDrawShadow.Move ctrlParent.Left + (4 * Screen.TwipsPerPixelX), _
                             ctrlParent.Top + (4 * Screen.TwipsPerPixelY), _
                             ctrlParent.Width, ctrlParent.Height
     ctrlDestDrawShadow.Visible = False
     DoEvents
     
     Set temp = CaptureForm(wForm)
     DoEvents
     
     ctrlDestDrawShadow.Picture = LoadPicture("")
     ctrlDestDrawShadow.PaintPicture temp, _
                                     -(ctrlFrame.Left + ctrlDestDrawShadow.Left + (wForm.Width - wForm.ScaleWidth) - 0), _
                                     -(ctrlFrame.Top + ctrlDestDrawShadow.Top + (wForm.Height - wForm.ScaleHeight) - 0)
     DoEvents
     
     ctrlDestDrawShadow.Visible = True
     
     Set temp = Nothing
     
     Call DrawMenuShadow(ctrlDestDrawShadow.hwnd, ctrlDestDrawShadow.hdc, Opacity)
End Sub

Private Sub DrawMenuShadow(ByVal hwnd As Long, _
                          ByVal hdc As Long, _
                          Optional ByVal Opacity As Long = 4)

  
     Dim Rec   As RECT
     Dim winW  As Long
     Dim winH  As Long
     
     Dim x     As Long
     Dim y     As Long
     Dim c     As Long
     Dim o     As Long
    
     GetWindowRect hwnd, Rec
     winW = (Rec.Right - Rec.Left)
     winH = (Rec.Bottom - Rec.Top)
        
     For o = 1 To Opacity
          For x = 1 To 4
               For y = 0 To 3
                    c = GetPixel(hdc, winW - x, y)
                    SetPixel hdc, winW - x, y, c
               Next y
     
               For y = 4 To 7
                    c = GetPixel(hdc, winW - x, y)
                    SetPixel hdc, winW - x, y, pMask(3 * x * (y - 3), c)
               Next y
     
               For y = 8 To winH - 5
                    c = GetPixel(hdc, winW - x, y)
                    SetPixel hdc, winW - x, y, pMask(15 * x, c)
               Next y
     
               For y = winH - 4 To winH - 1
                    c = GetPixel(hdc, winW - x, y)
                    SetPixel hdc, winW - x, y, pMask(3 * x * -(y - winH), c)
               Next y
          Next x
     
          For y = 1 To 4
               For x = 0 To 3
                    c = GetPixel(hdc, x, winH - y)
                    SetPixel hdc, x, winH - y, c
               Next x
     
               For x = 4 To 7
                    c = GetPixel(hdc, x, winH - y)
                    SetPixel hdc, x, winH - y, pMask(3 * (x - 3) * y, c)
               Next x
     
               For x = 8 To winW - 5
                    c = GetPixel(hdc, x, winH - y)
                    SetPixel hdc, x, winH - y, pMask(15 * y, c)
               Next x
          Next y
     Next o
End Sub

Private Function pMask(ByVal lScale As Long, _
                       ByVal lColor As Long) As Long

     Dim R        As Long
     Dim G        As Long
     Dim B        As Long
     Dim MyColour As Long

     MyColour = TranslateColour(lColor)
     R = MyColour And &HFF
     G = (MyColour And &HFF00&) \ &H100&
     B = (MyColour And &HFF0000) \ &H10000
     R = pTransform(lScale, R)
     G = pTransform(lScale, G)
     B = pTransform(lScale, B)
     pMask = RGB(R, G, B)
End Function

Private Function pTransform(ByVal lScale As Long, _
                            ByVal lColor As Long) As Long

    pTransform = lColor - Int(lColor * lScale / 255)
    ' - Function pTransform converts
    ' a RGB subcolor using a scale
    ' where 0 = 0 and 255 = lScale

End Function

Private Function TranslateColour(ByVal oClr As OLE_COLOR, _
                                 Optional hPal As Long = 0) As Long

     If OleTranslateColor(oClr, hPal, TranslateColour) Then
          TranslateColour = CLR_INVALID
     End If

End Function

Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long

    Dim Pic As PicBmp
    Dim IPic As IPicture ' IPicture requires a reference to "Standard OLE Types."
    Dim IID_IDispatch As GUID
    
    With IID_IDispatch ' Fill in with IDispatch Interface ID.
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    With Pic ' Fill Pic with necessary parts.
        .Size = Len(Pic)          ' Length of structure.
        .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
        .hBmp = hBmp              ' Handle to bitmap.
        .hPal = hPal              ' Handle to palette (may be null).
    End With

    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic) ' Create Picture object.
    
    Set CreateBitmapPicture = IPic ' Return the new Picture object.
End Function

Function CaptureWindow(ByVal hWndSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim R As Long
    Dim hdcSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
    
    hdcSrc = GetWindowDC(hWndSrc)                                ' Get device context for entire window.
    hDCMemory = CreateCompatibleDC(hdcSrc)                       ' Create a memory device context for the copy process.
    hBmp = CreateCompatibleBitmap(hdcSrc, WidthSrc, HeightSrc)   ' Create a bitmap and place it in the memory DC.
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    ' Get screen properties.
    RasterCapsScrn = GetDeviceCaps(hdcSrc, RASTERCAPS)           ' Raster capabilities.
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE               ' Palette support.
    PaletteSizeScrn = GetDeviceCaps(hdcSrc, SIZEPALETTE)         ' Size of palette.

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then           ' If the screen has a palette make a copy and realize it.
        ' Create a copy of the system palette.
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hdcSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)             ' Select the new palette into the memory DC and realize it.
        R = RealizePalette(hDCMemory)
    End If
    
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hdcSrc, LeftSrc, TopSrc, vbSrcCopy) ' Copy the on-screen image into the memory DC.
    hBmp = SelectObject(hDCMemory, hBmpPrev) ' Remove the new copy of the  on-screen image.
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then ' If the screen has a palette get back the palette that was selected in previously.
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    
    ' Release the device context resources back to the system.
    R = DeleteDC(hDCMemory)
    R = ReleaseDC(hWndSrc, hdcSrc)

    ' Call CreateBitmapPicture to create a picture object from the
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Function CaptureScreen() As Picture
    Dim hWndScreen As Long
    hWndScreen = GetDesktopWindow()
    ' Call CaptureWindow to capture the entire desktop give the handle
    ' and return the resulting Picture object.
    Set CaptureScreen = CaptureWindow(hWndScreen, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function

Function CaptureForm(frmsrc As Form) As Picture
    Dim hWndScreen As Long
    ' Call CaptureWindow to capture the entire desktop give the handle
    ' and return the resulting Picture object.
    Set CaptureForm = CaptureWindow(frmsrc.hwnd, 0, 0, frmsrc.ScaleX(frmsrc.Width, vbTwips, vbPixels), frmsrc.ScaleY(frmsrc.Height, vbTwips, vbPixels))
End Function

Function ConvertToPicture(ByVal Image As Long) As Picture
    Dim vIDispatch As GUID
    Dim vPic As PicBmp
    
    With vIDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With vPic
        .Size = Len(vPic)
        .Type = vbPicTypeBitmap
        .hBmp = Image
    End With
    Call OleCreatePictureIndirect(vPic, vIDispatch, 1, ConvertToPicture)
End Function
