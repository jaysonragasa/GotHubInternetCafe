VERSION 5.00
Begin VB.UserControl LabelControl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
End
Attribute VB_Name = "LabelControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."

Private Const CLR_INVALID                 As Integer = -1
Private Const DSS_NORMAL                  As Long = &H0
Private Const DST_COMPLEX                 As Long = &H0
Private Const DST_ICON                    As Long = &H3
Private Const DST_BITMAP                  As Long = &H4
Private Const DSS_DISABLED                As Long = &H20
Private Const DSS_MONO                    As Long = &H80

Private Type RECT
    Left                                    As Long
    Top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type

Public Enum ImageSizes
     img16 = 16
     Img24 = 24
     Img32 = 32
     Img48 = 48
     Img64 = 64
End Enum

'Property Variables:
Dim m_ImageSize As ImageSizes
Dim m_IconShadowed As Boolean
Dim m_ShadowOpacity As Integer
Dim m_PictureNormal As StdPicture
Dim m_Caption As String
'Dim m_ImageSize As ImageSizes


Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                    ByVal hBrush As Long, _
                                                                    ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, _
                                                                    ByVal wParam As Long, _
                                                                    ByVal n1 As Long, _
                                                                    ByVal n2 As Long, _
                                                                    ByVal n3 As Long, _
                                                                    ByVal n4 As Long, _
                                                                    ByVal un As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal crColor As Long) As Long
'Default Property Values:
Const m_def_ImageSize = 16


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
     BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
     UserControl.BackColor() = New_BackColor
     PropertyChanged "BackColor"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
     ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
     UserControl.ForeColor() = New_ForeColor
     PropertyChanged "ForeColor"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
     Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
     UserControl.Enabled() = New_Enabled
     PropertyChanged "Enabled"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
     Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
     Set UserControl.Font = New_Font
     PropertyChanged "Font"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
     BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
     UserControl.BackStyle() = New_BackStyle
     PropertyChanged "BackStyle"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
     BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
     UserControl.BorderStyle() = New_BorderStyle
     PropertyChanged "BorderStyle"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     UserControl.Refresh
End Sub

Private Sub UserControl_Click()
     RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
     RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
     BackColor = RGB(255, 0, 255)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
     RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
     RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
     Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
     Set UserControl.MouseIcon = New_MouseIcon
     PropertyChanged "MouseIcon"
     
     Call reDraw
End Property

Private Sub UserControl_OLECompleteDrag(Effect As Long)
     RaiseEvent OLECompleteDrag(Effect)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
     Set UserControl.Font = Ambient.Font
     m_Caption = "AdvanceLabel"
     Set m_PictureNormal = LoadPicture("")
     m_IconShadowed = True
     m_ShadowOpacity = 10
     m_ImageSize = m_def_ImageSize
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

     UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
     UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
     UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
     Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
     UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
     UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
     Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
     m_Caption = PropBag.ReadProperty("Caption", "AdvanceLabel")
     Set m_PictureNormal = PropBag.ReadProperty("PictureNormal", Nothing)
     m_IconShadowed = PropBag.ReadProperty("IconShadowed", m_IconShadowed)
     m_ShadowOpacity = PropBag.ReadProperty("ShadowOpacity", m_ShadowOpacity)
'     m_ImageSize = PropBag.ReadProperty("ImageSize", m_ImageSize)
     
     Call reDraw
     m_ImageSize = PropBag.ReadProperty("ImageSize", m_def_ImageSize)
End Sub

Private Sub UserControl_Resize()
     Call reDraw
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

     Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
     Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
     Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
     Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
     Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
     Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
     Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
     Call PropBag.WriteProperty("Caption", m_Caption, "AdvanceLabel")
     Call PropBag.WriteProperty("PictureNormal", m_PictureNormal, Nothing)
     Call PropBag.WriteProperty("IconShadowed", m_IconShadowed, m_IconShadowed)
     Call PropBag.WriteProperty("ShadowOpacity", m_ShadowOpacity, m_ShadowOpacity)
'     Call PropBag.WriteProperty("ImageSize", m_ImageSize, img16)
     Call PropBag.WriteProperty("ImageSize", m_ImageSize, m_def_ImageSize)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,AdvanceLabel
Public Property Get Caption() As String
     Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
     m_Caption = New_Caption
     PropertyChanged "Caption"
     
     Call reDraw
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureNormal() As Picture
     Set PictureNormal = m_PictureNormal
End Property

Public Property Set PictureNormal(ByVal New_PictureNormal As Picture)
     Set m_PictureNormal = New_PictureNormal
     PropertyChanged "PictureNormal"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get IconShadowed() As Boolean
     IconShadowed = m_IconShadowed
End Property

Public Property Let IconShadowed(ByVal New_IconShadowed As Boolean)
     m_IconShadowed = New_IconShadowed
     PropertyChanged "IconShadowed"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,10
Public Property Get ShadowOpacity() As Integer
     ShadowOpacity = m_ShadowOpacity
End Property

Public Property Let ShadowOpacity(ByVal New_ShadowOpacity As Integer)
     m_ShadowOpacity = New_ShadowOpacity
     PropertyChanged "ShadowOpacity"
     
     Call reDraw
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get ImageSize() As ImageSizes
     ImageSize = m_ImageSize
End Property

Public Property Let ImageSize(ByVal New_ImageSize As ImageSizes)
     m_ImageSize = New_ImageSize
     PropertyChanged "ImageSize"
     
     Call reDraw
End Property

Private Sub reDraw()
     UserControl.Cls
     
     Call DrawButtonImage(m_PictureNormal, _
                          0, 0, _
                          m_IconShadowed, _
                          UserControl.Enabled, _
                          m_ShadowOpacity)
                          
'     If m_PictureNormal <> Nothing Then
     UserControl.CurrentX = m_ImageSize + 4
     UserControl.CurrentY = (m_ImageSize - TextHeight("Yy")) / 2
     UserControl.Print m_Caption
'     End If
End Sub

Private Sub DrawButtonImage(ByRef m_Picture As StdPicture, _
                            ByVal x As Long, _
                            ByVal y As Long, _
                            ByVal bShadow As Boolean, _
                            ByVal Enabled As Boolean, _
                            ByVal ShadowOpacity As Integer)

  Dim lFlags As Long
  Dim hBrush As Long

    On Local Error Resume Next
    Select Case m_Picture.Type
     Case vbPicTypeBitmap
        lFlags = DST_BITMAP
     Case vbPicTypeIcon
        lFlags = DST_ICON
     Case Else
        lFlags = DST_COMPLEX
    End Select
    If bShadow Then
        'If m_OfficeXPStyle Then
            hBrush = CreateSolidBrush(BlendColour(vbHighlight, vbButtonShadow, m_ShadowOpacity))
         'Else 'M_OFFICEXPSTYLE = FALSE/0
            'hBrush = CreateSolidBrush(BlendColour(vbButtonShadow, BackColor, 60))
        'End If
    End If
    If Enabled Then
        DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x + 2, y + 2, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
        DrawState UserControl.hdc, IIf(False, hBrush, 0), 0, m_Picture.Handle, 0, x, y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or IIf(False, DSS_MONO, DSS_NORMAL)
     Else 'ENABLED = FALSE/0
        DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x + 2, y + 2, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or DSS_DISABLED
        DrawState UserControl.hdc, IIf(False, hBrush, 0), 0, m_Picture.Handle, 0, x, y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or DSS_DISABLED
    End If
    If bShadow Then
        DeleteObject hBrush
    End If

End Sub

Private Function BlendColour(ByVal oColorFrom As OLE_COLOR, _
                             ByVal oColorTo As OLE_COLOR, _
                             Optional ByVal alpha As Long = 128) As Long

  Dim lCFrom As Long
  Dim lCTo   As Long
  Dim lSrcR  As Long
  Dim lSrcG  As Long
  Dim lSrcB  As Long
  Dim lDstR  As Long
  Dim lDstG  As Long
  Dim lDstB  As Long

    On Local Error Resume Next
    lCFrom = TranslateColour(oColorFrom)
    lCTo = TranslateColour(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColour = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))

End Function

Private Function TranslateColour(ByVal oClr As OLE_COLOR, _
                                 Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColour) Then
        TranslateColour = CLR_INVALID
    End If

End Function

Private Sub DrawMenuShadow(ByVal hwnd As Long, _
                           ByVal hdc As Long, _
                           ByVal xOrg As Long, _
                           ByVal yOrg As Long)

  
  Dim Rec  As RECT
  Dim winW As Long
  Dim winH As Long

  Dim x    As Long
  Dim y    As Long
  Dim c    As Long
    
    GetWindowRect hwnd, Rec
    winW = (Rec.Right - Rec.Left)
    winH = (Rec.Bottom - Rec.Top)
    c = TranslateColour(UserControl.BackColor)
    For x = 1 To 4
        For y = 0 To 3
            SetPixel hdc, winW - x, y, c
        Next y
        For y = 4 To 7
            SetPixel hdc, winW - x, y, pMask(3 * x * (y - 3), c)
        Next y
        For y = 8 To winH - 5
            SetPixel hdc, winW - x, y, pMask(15 * x, c)
        Next y
        For y = winH - 4 To winH - 1
            SetPixel hdc, winW - x, y, pMask(3 * x * -(y - winH), c)
        Next y
    Next x
    For y = 1 To 4
        For x = 0 To 3
            SetPixel hdc, x, winH - y, c
        Next x
        For x = 4 To 7
            SetPixel hdc, x, winH - y, pMask(3 * (x - 3) * y, c)
        Next x
        For x = 8 To winW - 5
            SetPixel hdc, x, winH - y, pMask(15 * y, c)
        Next x
    Next y

End Sub

Private Function pMask(ByVal lScale As Long, _
                       ByVal lColor As Long) As Long

  Dim R        As Long
  Dim g        As Long
  Dim b        As Long
  Dim MyColour As Long

    MyColour = TranslateColour(lColor)
    R = MyColour And &HFF
    g = (MyColour And &HFF00&) \ &H100&
    b = (MyColour And &HFF0000) \ &H10000
    R = pTransform(lScale, R)
    g = pTransform(lScale, g)
    b = pTransform(lScale, b)
    pMask = RGB(R, g, b)

End Function

Private Function pTransform(ByVal lScale As Long, _
                            ByVal lColor As Long) As Long

    pTransform = lColor - Int(lColor * lScale / 255)
    ' - Function pTransform converts
    ' a RGB subcolor using a scale
    ' where 0 = 0 and 255 = lScale

End Function

