VERSION 5.00
Begin VB.UserControl AdvanceLabelControl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "AdvanceLabelControl.ctx":0000
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   166
End
Attribute VB_Name = "AdvanceLabelControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

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
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal crColor As Long) As Long
                                                                                              
Private Const CLR_INVALID                    As Integer = -1
Private Const DSS_NORMAL                     As Long = &H0
Private Const DST_COMPLEX                    As Long = &H0
Private Const DST_ICON                       As Long = &H3
Private Const DST_BITMAP                     As Long = &H4
Private Const DSS_DISABLED                   As Long = &H20
Private Const DSS_MONO                       As Long = &H80

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
    Left                                     As Long
    Top                                      As Long
    Right                                    As Long
    Bottom                                   As Long
End Type

Public Enum ImageSizes
     img16 = 16
     Img24 = 24
     Img32 = 32
     Img48 = 48
     Img64 = 64
End Enum

Public Enum OnHoverStyles
     Underlined = 0
     Bolded = 1
     Italics = 2
     BoldAndUnderline = 3
End Enum

Private WithEvents MouseTimer                As ccrpTimer
Attribute MouseTimer.VB_VarHelpID = -1
Dim CurPos                                   As POINTAPI
Dim CurHWND                                  As Long

'Property Variables:
Dim m_OnHoverIconShadow As Boolean
Dim m_OnHoverForeColor                       As OLE_COLOR
Dim m_UseHandPointer                         As Boolean
Dim m_OnHoverStyle                           As OnHoverStyles
Dim m_AutoResize                             As Boolean
Dim m_HasIcon                                As Boolean
Dim m_NormalIcon                             As StdPicture
Dim m_IconSize                               As ImageSizes
Dim m_Caption                                As String
Dim m_ShadowOpacity                          As Integer

'Default Property Values:
Const m_def_OnHoverIconShadow = True
Const m_def_OnHoverForeColor = &H80000012
Const m_def_UseHandPointer = False
Const m_def_OnHoverStyle = 0
Const m_def_AutoResize = False
Const m_def_HasIcon = False
Const m_def_NormalIcon = 0
Const m_def_IconSize = 16
Const m_def_Caption = "AdvanceLabel 1.0"
Const m_def_ShadowOpacity = 10

'Event Declarations:
Event Click()                                                                   'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick()                                                                'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer)                             'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer)                                             'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer)                               'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)  'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)  'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)    'MappingInfo=UserControl,UserControl,-1,MouseUp




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
          hBrush = CreateSolidBrush(BlendColour(vbHighlight, vbButtonShadow, m_ShadowOpacity))
        
        'If m_OfficeXPStyle Then
            'hBrush = CreateSolidBrush(BlendColour(vbHighlight, vbButtonShadow, 10))
         'Else 'M_OFFICEXPSTYLE = FALSE/0
            'hBrush = CreateSolidBrush(BlendColour(vbButtonShadow, BackColor, 60))
        'End If
     End If
     
     If Enabled Then
          DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x, y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
     Else 'ENABLED = FALSE/0
          DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x, y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or DSS_DISABLED
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
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
     Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
     Set UserControl.MouseIcon = New_MouseIcon
     PropertyChanged "MouseIcon"
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
     
     If MouseTimer.Enabled = False Then MouseTimer.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,AdvanceLabel 1.0
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
'MemberInfo=0,0,0,True
Public Property Get HasIcon() As Boolean
     HasIcon = m_HasIcon
End Property

Public Property Let HasIcon(ByVal New_HasIcon As Boolean)
     m_HasIcon = New_HasIcon
     PropertyChanged "HasIcon"
     
     Call reDraw
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get NormalIcon() As StdPicture
     Set NormalIcon = m_NormalIcon
End Property

Public Property Set NormalIcon(ByVal New_NormalIcon As StdPicture)
     Set m_NormalIcon = New_NormalIcon
     PropertyChanged "NormalIcon"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=22,0,0,16
Public Property Get IconSize() As ImageSizes
     IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_IconSize As ImageSizes)
     m_IconSize = New_IconSize
     PropertyChanged "IconSize"
     
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get AutoResize() As Boolean
     AutoResize = m_AutoResize
End Property

Public Property Let AutoResize(ByVal New_AutoResize As Boolean)
     m_AutoResize = New_AutoResize
     PropertyChanged "AutoResize"
     
     Call reDraw
     Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,0
Public Property Get OnHoverStyle() As OnHoverStyles
     OnHoverStyle = m_OnHoverStyle
End Property

Public Property Let OnHoverStyle(ByVal New_OnHoverStyle As OnHoverStyles)
     m_OnHoverStyle = New_OnHoverStyle
     PropertyChanged "OnHoverStyle"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H80000000&
Public Property Get OnHoverForeColor() As OLE_COLOR
     OnHoverForeColor = m_OnHoverForeColor
End Property

Public Property Let OnHoverForeColor(ByVal New_OnHoverForeColor As OLE_COLOR)
     m_OnHoverForeColor = New_OnHoverForeColor
     PropertyChanged "OnHoverForeColor"
     
     Call reDraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get UseHandPointer() As Boolean
     UseHandPointer = m_UseHandPointer
End Property

Public Property Let UseHandPointer(ByVal New_UseHandPointer As Boolean)
     m_UseHandPointer = New_UseHandPointer
     PropertyChanged "UseHandPointer"
     
     If m_UseHandPointer Then
          UserControl.MousePointer = MousePointerConstants.vbCustom
     Else: UserControl.MousePointer = MousePointerConstants.vbArrow
     End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get OnHoverIconShadow() As Boolean
     OnHoverIconShadow = m_OnHoverIconShadow
End Property

Public Property Let OnHoverIconShadow(ByVal New_OnHoverIconShadow As Boolean)
     m_OnHoverIconShadow = New_OnHoverIconShadow
     PropertyChanged "OnHoverIconShadow"
     
     Call reDraw
End Property

Private Sub UserControl_Initialize()
     Set MouseTimer = New ccrpTimer
     MouseTimer.Interval = 10
End Sub

'
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
     Set UserControl.Font = Ambient.Font
     m_Caption = m_def_Caption
     m_ShadowOpacity = m_def_ShadowOpacity
     Set m_NormalIcon = Nothing
     m_IconSize = m_def_IconSize
     m_HasIcon = m_def_HasIcon
     m_AutoResize = m_def_AutoResize
     m_OnHoverStyle = m_def_OnHoverStyle
     m_UseHandPointer = m_def_UseHandPointer
     m_OnHoverForeColor = m_def_OnHoverForeColor
     m_OnHoverIconShadow = m_def_OnHoverIconShadow
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

     UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
     UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
     UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
     Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
     m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
     m_ShadowOpacity = PropBag.ReadProperty("ShadowOpacity", m_def_ShadowOpacity)
     Set m_NormalIcon = PropBag.ReadProperty("NormalIcon", Nothing)
     m_IconSize = PropBag.ReadProperty("IconSize", m_def_IconSize)
     m_HasIcon = PropBag.ReadProperty("HasIcon", m_def_HasIcon)
     Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
     m_AutoResize = PropBag.ReadProperty("AutoResize", m_def_AutoResize)
     m_OnHoverStyle = PropBag.ReadProperty("OnHoverStyle", m_def_OnHoverStyle)
     m_UseHandPointer = PropBag.ReadProperty("UseHandPointer", m_def_UseHandPointer)
     Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
     m_OnHoverForeColor = PropBag.ReadProperty("OnHoverForeColor", m_def_OnHoverForeColor)
     
     Call reDraw
     
     m_OnHoverIconShadow = PropBag.ReadProperty("OnHoverIconShadow", m_def_OnHoverIconShadow)
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
     Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
     Call PropBag.WriteProperty("ShadowOpacity", m_ShadowOpacity, m_def_ShadowOpacity)
     Call PropBag.WriteProperty("NormalIcon", m_NormalIcon, m_def_NormalIcon)
     Call PropBag.WriteProperty("IconSize", m_IconSize, m_def_IconSize)
     Call PropBag.WriteProperty("HasIcon", m_HasIcon, m_def_HasIcon)
     Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
     Call PropBag.WriteProperty("AutoResize", m_AutoResize, m_def_AutoResize)
     Call PropBag.WriteProperty("OnHoverStyle", m_OnHoverStyle, m_def_OnHoverStyle)
     Call PropBag.WriteProperty("UseHandPointer", m_UseHandPointer, m_def_UseHandPointer)
     Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
     Call PropBag.WriteProperty("OnHoverForeColor", m_OnHoverForeColor, m_def_OnHoverForeColor)
     Call PropBag.WriteProperty("OnHoverIconShadow", m_OnHoverIconShadow, m_def_OnHoverIconShadow)
End Sub

Private Sub reDraw()
     Dim cX As Integer
     Dim cY As Integer
     
     Cls
     
     If m_HasIcon Then
          If m_OnHoverIconShadow And MouseTimer.Enabled = True Then
               Call DrawButtonImage(m_NormalIcon, 2, 2, m_OnHoverIconShadow, Enabled, m_ShadowOpacity)
               Call DrawButtonImage(m_NormalIcon, 0, 0, False, Enabled, m_ShadowOpacity)
          Else
               Call DrawButtonImage(m_NormalIcon, 0, 0, False, Enabled, m_ShadowOpacity)
          End If
          
          cX = m_IconSize + 4
          cY = (m_IconSize - TextHeight(m_Caption)) / 2
     Else
          cX = 1
          cY = (ScaleHeight - TextHeight(m_Caption)) / 2
     End If
     
     If m_AutoResize Then
          If m_HasIcon Then
               If m_OnHoverIconShadow Then
                    Height = (m_IconSize + 2) * Screen.TwipsPerPixelY
                    Width = (m_IconSize + 4 + TextWidth(m_Caption)) * Screen.TwipsPerPixelX
               Else
                    Height = m_IconSize * Screen.TwipsPerPixelY
                    Width = (m_IconSize + 4 + TextWidth(m_Caption)) * Screen.TwipsPerPixelX
               End If
          Else
               Height = TextHeight(m_Caption) * Screen.TwipsPerPixelY
               Width = TextWidth(m_Caption) * Screen.TwipsPerPixelX
          End If
     End If
     
     CurrentX = cX: CurrentY = cY
     Print m_Caption
End Sub

Private Sub MouseTimer_Timer(ByVal Milliseconds As Long)
     Dim nHwnd      As Long
     
     GetCursorPos CurPos
     
     nHwnd = WindowFromPoint(CurPos.x, CurPos.y)
     
     If nHwnd = hwnd Then
          If m_OnHoverStyle = Underlined And Font.Underline = False Then Font.Underline = True
          If m_OnHoverStyle = Bolded And Font.Bold = False Then Font.Bold = True
          If m_OnHoverStyle = Italics And Font.Italic = False Then Font.Italic = True
          If m_OnHoverStyle = BoldAndUnderline And Font.Bold = False And Font.Underline = False Then Font.Bold = True: Font.Underline = True
          
          If m_OnHoverIconShadow = False Then m_OnHoverIconShadow = True
          
          ForeColor = m_OnHoverForeColor
     Else
          If Font.Underline = True And m_OnHoverStyle = Underlined Then Font.Underline = False
          If Font.Bold = True And m_OnHoverStyle = Bolded Then Font.Bold = False
          If Font.Italic = True And m_OnHoverStyle = Italics Then Font.Italics = False
          
          If m_OnHoverIconShadow = True Then m_OnHoverIconShadow = False
          
          ForeColor = m_def_OnHoverForeColor
          
          MouseTimer.Enabled = False
     End If
     
     Call reDraw
End Sub

