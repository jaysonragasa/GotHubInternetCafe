VERSION 5.00
Begin VB.UserControl InfoHeader 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
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
   PropertyPages   =   "InfoHeader.ctx":0000
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   Begin VB.PictureBox picGrad 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   -15
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   0
      Top             =   0
      Width           =   1785
      Begin VB.Image imgClickMe 
         Height          =   270
         Left            =   1140
         Top             =   75
         Width           =   555
      End
      Begin VB.Image imgIco 
         Height          =   240
         Left            =   120
         Picture         =   "InfoHeader.ctx":0028
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   165
         Width           =   465
      End
   End
End
Attribute VB_Name = "InfoHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' I created my own
' Information Header v1.0
' Coded By: Jayson Ragasa
' Copyright© 2004 Baguio City, Philippines
' ----------------------------------------
'
' Information Header is a customizable object that you can use in your programs.
'
'
' -----------------------------------------------------------------------------------------------------------
'    Information Header v1.0 (released) Wednesday, April 21, 2004
' -----------------------------------------------------------------------------------------------------------
'         * Caption
'         * LeftColor And RightColor - Customizable Gradient Colors
'         * MaxFill
'         * Horizontaol And Vertical gradient style
'         * HasIcon - left icon position. or you can browse BMP, GIF, JPG any picture you want
'         * Left, Center, Right, IconBottm caption alignments.
'         * ForeColor - choose your forecolor style
'         * Font - choose you font style.
'         * MultiLine - multi line supported. usr pipe "|" as your vbCrLf.
' -----------------------------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------------------------
'    Information Header v2.0
' -----------------------------------------------------------------------------------------------------------
'         * EdgeRadius - now you can set the edge radius to make it round
'         * RoundEdgeStyle - you can set the rounded edge position to TopLeft, TopRight, BottomLeft, _
'                          - BottomRight, Left, Right, Top, Bottom, Whole, or you can set it to None.
'         * InfoHeaderStyle - you can make just a header or you can set to frame like (Container)
'         * HeaderHeight - Header height.
'         * BackColor - if InfoHeaderStyle is Frame.. you can set its BackColor or any color you want.
'         * BorderColor - If InfoHeaderStyle is Frame.. you can set its BorderColor or any color you want.
'
' any comments or suggestion is appriciated
' if you like this, please vote for it! tnx!


Option Explicit

Public Enum InfoHeaderStyl
     Header = 0
     Frame = 1
End Enum

Public Enum CaptionAlignment
     AlignLeft = 0
     AlignCenter = 1
     AlignRight = 2
     AlignIconBottom = 3
End Enum

Public Enum GradStyle
     GradientHorizontal = 0
     GradientVertical = 1
End Enum

Public Enum RoundEdgeStyl
     REP_TopLeft = 0
     REP_TopRight = 1
     REP_BottomLeft = 2
     REP_BottomRight = 3
     REP_Left = 4
     REP_Right = 5
     REP_Top = 6
     REP_Bottom = 7
     REP_Whole = 8
     REP_None = 9
End Enum

Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    alpha As Integer
End Type
    
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Dim IHS                  As InfoHeaderStyl
Dim HeaderSize           As Integer
Dim Fra_BackColor        As OLE_COLOR
Dim Fra_BorderColor      As OLE_COLOR

Dim CaptionAlign         As CaptionAlignment
Dim MltiLyn              As Boolean

Dim b_HasIcon            As Boolean
Dim pictIcon             As StdPicture

Dim GradStyl             As GradStyle
Dim Color_LEFT           As OLE_COLOR
Dim Color_RIGHT          As OLE_COLOR
Dim iMaxFill             As Integer

Dim RadiusSize           As Integer
Dim res                  As RoundEdgeStyl

Dim iMargin              As Integer
Dim CurWidth             As Long

Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V  As Long = &H1

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Dim rRed As Long, rBlue As Long, rGreen As Long

Dim R1 As Integer, R2 As Integer
Dim B1 As Integer, B2 As Integer
Dim G1 As Integer, G2 As Integer

Public Event Click()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Property Let InfoHeaderStyle(ByVal newVal As InfoHeaderStyl)
     IHS = newVal
     PropertyChanged "InfoHeaderStyle"
     
     Call UserControl_Resize
     'Call reDraw(True)
End Property
Public Property Get InfoHeaderStyle() As InfoHeaderStyl
     InfoHeaderStyle = IHS
End Property

Public Property Let HeaderHeight(ByVal newVal As Integer)
     HeaderSize = newVal
     PropertyChanged "HeaderHeight"
     
     Call UserControl_Resize
     'Call reDraw(True)
End Property
Public Property Get HeaderHeight() As Integer
     HeaderHeight = HeaderSize
End Property

Public Property Let FrameBackColor(ByVal newVal As OLE_COLOR)
     Fra_BackColor = newVal
     PropertyChanged "FrameBackColor"
     
     Call reDraw(True)
End Property
Public Property Get FrameBackColor() As OLE_COLOR
     FrameBackColor = Fra_BackColor
End Property

Public Property Let FrameBorderColor(ByVal newVal As OLE_COLOR)
     Fra_BorderColor = newVal
     PropertyChanged "FrameBorderColor"
     
     Call reDraw(True)
End Property
Public Property Get FrameBorderColor() As OLE_COLOR
     FrameBorderColor = Fra_BorderColor
End Property

Public Property Let GradientStyle(ByVal newVal As GradStyle)
     GradStyl = newVal
     PropertyChanged "GradientStyle"
     
     Call reDraw(True)
End Property
Public Property Get GradientStyle() As GradStyle
     GradientStyle = GradStyl
End Property

Public Property Let LeftColor(ByVal newVal As OLE_COLOR)
     Color_LEFT = ConvertRGBFormat(newVal)
     PropertyChanged "LeftColor"
     
     Call reDraw(True)
End Property
Public Property Get LeftColor() As OLE_COLOR
     LeftColor = Color_LEFT
End Property

Public Property Let RightColor(ByVal newVal As OLE_COLOR)
     Color_RIGHT = ConvertRGBFormat(newVal)
     PropertyChanged "RightColor"
     
     Call reDraw(True)
End Property
Public Property Get RightColor() As OLE_COLOR
     RightColor = Color_RIGHT
End Property

Public Property Set FontName(ByVal newVal As StdFont)
     Set lblCaption.Font = newVal
     PropertyChanged "FontName"
     
     Call reDraw(False)
End Property
Public Property Get FontName() As StdFont
     Set FontName = lblCaption.Font
End Property

Public Property Let ForeColor(ByVal newVal As OLE_COLOR)
     lblCaption.ForeColor = newVal
     PropertyChanged "ForeColor"

     Call reDraw(False)
End Property
Public Property Get ForeColor() As OLE_COLOR
     ForeColor = lblCaption.ForeColor
End Property

Public Property Let Margin(ByVal newVal As Integer)
     iMargin = newVal
     PropertyChanged "Margin"
     
     Call reDraw(False)
End Property
Public Property Get Margin() As Integer
     Margin = iMargin
End Property

Public Property Let Caption(ByVal newVal As String)
     If Not MltiLyn Then
          lblCaption.Caption = newVal
     ElseIf MltiLyn Then
          lblCaption.Caption = DoMultiLining(newVal)
     End If
     
     PropertyChanged "Caption"
     
     Call reDraw(False)
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "TextStyle"
     Caption = lblCaption.Caption
End Property

Public Property Let Alignment(ByVal newVal As CaptionAlignment)
     If HasIcon = False And newVal = AlignIconBottom Then
          MsgBox "Cannot set alignment because 'HasIcon' is false"
          
          Exit Property
     End If
     
     CaptionAlign = newVal
     PropertyChanged "Alignment"
     
     Call reDraw(False)
End Property
Public Property Get Alignment() As CaptionAlignment
     Alignment = CaptionAlign
End Property

Public Property Let MultiLine(ByVal newVal As Boolean)
     MltiLyn = newVal
     
     If newVal = True Then
          lblCaption.Caption = DoMultiLining(Caption)
     ElseIf newVal = False Then
          lblCaption.Caption = RemoveCRLF(Caption)
     End If
     
     PropertyChanged "MultiLine"
     
     Call reDraw(False)
End Property
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_ProcData.VB_Invoke_Property = "TextStyle"
     MultiLine = MltiLyn
End Property

Public Property Let HasIcon(ByVal newVal As Boolean)
     If newVal = False Then
          If Alignment = AlignIconBottom Then
               Alignment = AlignLeft
               CaptionAlign = AlignLeft
          End If
     End If
     b_HasIcon = newVal
     PropertyChanged "HasIcon"
     
     Call reDraw(False)
End Property
Public Property Get HasIcon() As Boolean
     HasIcon = b_HasIcon
End Property

Public Property Set Picture(ByVal newVal As StdPicture)
     Set imgIco.Picture = newVal
     PropertyChanged "Picture"
     
     Call reDraw(False)
End Property
Public Property Get Picture() As StdPicture
     Set Picture = imgIco.Picture
End Property

Public Property Let EdgeRadiusSize(ByVal newVal As Integer)
     RadiusSize = newVal
     PropertyChanged "EdgeRadiusSize"
     
     Call reDraw(False)
End Property
Public Property Get EdgeRadiusSize() As Integer
     EdgeRadiusSize = RadiusSize
End Property

Public Property Let RoundEdgePosition(ByVal newVal As RoundEdgeStyl)
     res = newVal
     PropertyChanged "RoundEdgeStyle"
     
     Call reDraw(False)
End Property
Public Property Get RoundEdgePosition() As RoundEdgeStyl
     RoundEdgePosition = res
End Property

Public Function l_hDC() As Long
     l_hDC = hdc
     'UserControl.hDC = l_hDC
End Function

Private Sub imgClickMe_Click()
     RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
     IHS = Header
     HeaderSize = 26
     
     Fra_BackColor = ConvertRGBFormat(SystemColorConstants.vbWindowBackground)
     Fra_BorderColor = ConvertRGBFormat(SystemColorConstants.vb3DShadow)
     
     GradStyl = GradientHorizontal
     Color_LEFT = ConvertRGBFormat(SystemColorConstants.vbActiveTitleBar)
     Color_RIGHT = ConvertRGBFormat(SystemColorConstants.vb3DFace)
     iMaxFill = 100
     
     lblCaption.Font = "Tahoma"
     lblCaption.ForeColor = ConvertRGBFormat(SystemColorConstants.vbActiveTitleBarText)
     lblCaption.FontBold = True
     lblCaption.Caption = "Information Header v2.0"
     
     CaptionAlign = AlignLeft
     MltiLyn = False
     
     b_HasIcon = True
     
     RadiusSize = 0
     res = REP_None
     
     iMargin = 4
     
     Call reDraw(True)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     IHS = PropBag.ReadProperty("InfoHeaderStyle", InfoHeaderStyl.Header)
     HeaderSize = PropBag.ReadProperty("HeaderHeight", 26)
     Fra_BackColor = PropBag.ReadProperty("FrameBackColor", ConvertRGBFormat(SystemColorConstants.vbWindowBackground))
     Fra_BorderColor = PropBag.ReadProperty("FrameBorderColor", ConvertRGBFormat(SystemColorConstants.vb3DShadow))
     
     GradStyl = PropBag.ReadProperty("GradientStyle", GradStyle.GradientHorizontal)
     Color_LEFT = PropBag.ReadProperty("LeftColor", ConvertRGBFormat(SystemColorConstants.vb3DShadow))
     Color_RIGHT = PropBag.ReadProperty("RightColor", ConvertRGBFormat(SystemColorConstants.vb3DFace))
     
     iMargin = PropBag.ReadProperty("Margin", 4)
     
     lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", ConvertRGBFormat(SystemColorConstants.vb3DDKShadow))
     Set lblCaption.Font = PropBag.ReadProperty("FontName", "Tahoma")
     lblCaption.Caption = PropBag.ReadProperty("Caption", "Information Header v1.0")
     MltiLyn = PropBag.ReadProperty("MultiLine", False)
     CaptionAlign = PropBag.ReadProperty("Alignment", CaptionAlignment.AlignLeft)
     
     b_HasIcon = PropBag.ReadProperty("HasIcon", True)
     Set imgIco.Picture = PropBag.ReadProperty("Picture", imgIco.Picture)
     
     RadiusSize = PropBag.ReadProperty("EdgeRadiusSize", 0)
     res = PropBag.ReadProperty("RoundEdgeSTyle", REP_TopLeft)
     
     Call UserControl_Resize
     Call reDraw(True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     PropBag.WriteProperty "InfoHeaderStyle", IHS
     PropBag.WriteProperty "HeaderHeight", HeaderSize
     PropBag.WriteProperty "FrameBackColor", Fra_BackColor
     PropBag.WriteProperty "FrameBorderColor", Fra_BorderColor
     
     PropBag.WriteProperty "GradientStyle", GradStyl
     PropBag.WriteProperty "LeftColor", Color_LEFT
     PropBag.WriteProperty "RightColor", Color_RIGHT
     
     PropBag.WriteProperty "Margin", iMargin
     
     PropBag.WriteProperty "FontName", lblCaption.Font
     PropBag.WriteProperty "ForeColor", lblCaption.ForeColor
     PropBag.WriteProperty "Caption", lblCaption.Caption
     PropBag.WriteProperty "MultiLine", MltiLyn
     PropBag.WriteProperty "Alignment", CaptionAlign
     
     PropBag.WriteProperty "HasIcon", b_HasIcon
     PropBag.WriteProperty "Picture", imgIco.Picture
     
     PropBag.WriteProperty "EdgeRadiusSize", RadiusSize
     PropBag.WriteProperty "RoundEdgeStyle", res
End Sub

Sub reDraw(ByVal reDrawGrad As Boolean)
     Dim tmpTop     As Long
     Dim tmpHeight  As Long
     Dim tmpRGN     As Long
     
     UserControl.Cls
     BackColor = Fra_BackColor
     Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), Fra_BorderColor, B
     imgClickMe.Move 0, 0, picGrad.ScaleWidth, picGrad.ScaleHeight
     
     If reDrawGrad = True Then
          picGrad.Cls
          picGrad.BackColor = Fra_BackColor
          Gradient picGrad, Color_LEFT, Color_RIGHT, GradStyl
     End If
     
     If HasIcon Then
          imgIco.Visible = True
          
          If MltiLyn = False Then
               lblCaption.Top = (picGrad.Height - lblCaption.Height) \ 2
               
               imgIco.Move tX(iMargin), (picGrad.Height - imgIco.Height) \ 2
          ElseIf MltiLyn = True Then
               imgIco.Left = tX(iMargin)
               
               If imgIco.Height > lblCaption.Height Then
                    imgIco.Top = (picGrad.Height - imgIco.Height) \ 2
                    
                    lblCaption.Top = imgIco.Top
               Else
                    lblCaption.Top = (picGrad.Height - lblCaption.Height) \ 2
                    
                    imgIco.Top = lblCaption.Top
               End If
          End If
          
          If CaptionAlign = AlignLeft Then
               lblCaption.Left = imgIco.Left + imgIco.Width + tX(7)
               
               lblCaption.Alignment = AlignmentConstants.vbLeftJustify
          ElseIf CaptionAlign = AlignCenter Then
               lblCaption.Alignment = AlignmentConstants.vbCenter
               
               lblCaption.Left = (((picGrad.Width - (imgIco.Left + imgIco.Width)) - lblCaption.Width) \ 2) + (imgIco.Left + imgIco.Width)
          ElseIf CaptionAlign = AlignRight Then
               lblCaption.Left = picGrad.ScaleWidth - lblCaption.Width - tX(iMargin)
               
               lblCaption.Alignment = AlignmentConstants.vbRightJustify
          ElseIf CaptionAlign = AlignIconBottom Then
               lblCaption.Alignment = AlignmentConstants.vbLeftJustify
               tmpHeight = imgIco.Height + lblCaption.Height + tY(2)
               tmpTop = (picGrad.ScaleHeight - tmpHeight) \ 2
               
               imgIco.Top = tmpTop
               
               lblCaption.Top = imgIco.Top + imgIco.Height + tY(2)
               lblCaption.Left = imgIco.Left
          End If
     ElseIf Not HasIcon Then
          imgIco.Visible = False
          
          lblCaption.Top = (picGrad.Height - lblCaption.Height) \ 2
          
          If CaptionAlign = AlignLeft Then
               lblCaption.Left = tX(4)
               
               lblCaption.Alignment = AlignmentConstants.vbLeftJustify
          ElseIf CaptionAlign = AlignCenter Then
               lblCaption.Left = ((picGrad.ScaleWidth - lblCaption.Width) \ 2)
               
               lblCaption.Alignment = AlignmentConstants.vbCenter
          ElseIf CaptionAlign = AlignRight Then
               lblCaption.Left = picGrad.ScaleWidth - lblCaption.Width - tX(iMargin)
               
               lblCaption.Alignment = AlignmentConstants.vbRightJustify
          End If
     End If
     
     If res = REP_TopLeft Then
          tmpRGN = CreateRoundRectRgn(0, 0, ScaleWidth + RadiusSize, ScaleHeight + RadiusSize, RadiusSize, RadiusSize)
          
     ElseIf res = REP_TopRight Then
          tmpRGN = CreateRoundRectRgn(0 - RadiusSize, 0, ScaleWidth + 1, ScaleHeight + RadiusSize, RadiusSize, RadiusSize)
          
     ElseIf res = REP_BottomLeft Then
          tmpRGN = CreateRoundRectRgn(0, 0 - RadiusSize, ScaleWidth + RadiusSize, ScaleHeight + 1, RadiusSize, RadiusSize)
          
     ElseIf res = REP_BottomRight Then
          tmpRGN = CreateRoundRectRgn(0 - RadiusSize, 0 - RadiusSize, ScaleWidth + 1, ScaleHeight + 1, RadiusSize, RadiusSize)
          
     ElseIf res = REP_Top Then
          tmpRGN = CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + RadiusSize, RadiusSize, RadiusSize)
          
     ElseIf res = REP_Bottom Then
          tmpRGN = CreateRoundRectRgn(0, 0 - RadiusSize, ScaleWidth + 1, ScaleHeight + 1, RadiusSize, RadiusSize)
          
     ElseIf res = REP_Left Then
          tmpRGN = CreateRoundRectRgn(0, 0, ScaleWidth + RadiusSize, ScaleHeight + 1, RadiusSize, RadiusSize)
          
     ElseIf res = REP_Right Then
          tmpRGN = CreateRoundRectRgn(0 - RadiusSize, 0, ScaleWidth + 1, ScaleHeight + 1, RadiusSize, RadiusSize)
     
     ElseIf res = REP_Whole Then
          tmpRGN = CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1, RadiusSize, RadiusSize)
          
     ElseIf res = REP_None Then
          tmpRGN = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 0, 0)
     End If
     
     SetWindowRgn hwnd, tmpRGN, True
End Sub

Private Sub UserControl_Resize()
     'If Height < 510 Then Height = 510
     
     If IHS = Header Then
          picGrad.Move 0, 0, ScaleWidth, ScaleHeight
     ElseIf IHS = Frame Then
          picGrad.Move 0, 0, ScaleWidth, HeaderSize
     End If
     
     If ScaleWidth <> CurWidth Then
          Call reDraw(True)
     Else
          Call reDraw(False)
     End If
     
     If IHS = Header Then Call reDraw(True)
     
     CurWidth = ScaleWidth
End Sub

' DoMultiLining v1.0
Private Function DoMultiLining(ByVal sText As String) As String
     Dim exLines    As Variant
     Dim i          As Integer
     Dim tmp        As String
     
     If InStr(1, sText, "|") <> 0 Then
          exLines = Split(sText, "|")
               
          For i = 0 To UBound(exLines)
               tmp = tmp + exLines(i) + vbCrLf
          Next i
          
          DoMultiLining = Left$(tmp, Len(tmp) - 2)
     Else
          DoMultiLining = sText
     End If
End Function

Private Function RemoveCRLF(ByVal sText As String) As String
     Dim i     As Integer
     Dim c     As String
     Dim ret   As String
     
     For i = 1 To Len(sText)
          c = Mid$(sText, i, 1)
          
          If Asc(c) = 10 Then
          ElseIf Asc(c) = 13 Then
               ret = ret + "|"
          Else
               ret = ret + c
          End If
     Next i
     
     RemoveCRLF = ret
End Function

Private Function tX(ByVal pX As Integer) As Integer
     'tX = pX * Screen.TwipsPerPixelX
     tX = pX
End Function

Private Function tY(ByVal pY As Integer) As Integer
     'tY = pY * Screen.TwipsPerPixelY
     tY = pY
End Function
'
'Private Function LongToUShort(ULong As Long) As Integer
'   LongToUShort = CInt(ULong - &H10000)
'End Function
'
'Private Function UShortToLong(Ushort As Integer) As Long
'   UShortToLong = (CLng(Ushort) And &HFFFF&)
'End Function

'Private Sub Gradient(ByRef picBox, _
'                     ByVal StartColor As Long, _
'                     ByVal EndColor As Long, _
'                     ByVal GradType As GradStyle)
'
'     Dim Vert(1) As TRIVERTEX
'     Dim gRect As GRADIENT_RECT
'
'     picBox.AutoRedraw = True
'     picBox.ScaleMode = vbPixels
'
'     Call SetupColors(StartColor, EndColor): DoEvents
'     With Vert(0)
'          .X = 0
'          .Y = 0
'          .Red = CInt(R1) And &H0&
'          .Green = CInt(G1) And &H0&
'          .Blue = CInt(B1) And &H0&
'          .alpha = 0&
'     End With
'
'     With Vert(1)
'          .X = picBox.ScaleWidth
'          .Y = picBox.ScaleHeight
'          .Red = CInt(R2)
'          .Green = CInt(G2) 'LongToUShort(CLng(G2))
'          .Blue = CInt(B2) ' LongToUShort(CLng(B2))
'          .alpha = 0&
'     End With
'
'     gRect.UpperLeft = 1
'     gRect.LowerRight = 0
'     'replace GRADIENT_FILL_RECT_H with GRADIENT_FILL_RECT_V  to paint
'     'the form with vertically gradient, instead of horizontally gradient
'     If GradType = GradientVertical Then
'          GradientFillRect picBox.hdc, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V
'     ElseIf GradType = GradientHorizontal Then
'          GradientFillRect picBox.hdc, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H
'     End If
'End Sub
Private Sub Gradient(ByRef picBox, _
                    ByVal StartColor As Long, _
                    ByVal EndColor As Long, _
                    ByVal GradType As GradStyle)

     Dim i     As Integer
     Dim Color As Long

     Dim Size  As Long

     Dim GradR As Integer, GradB As Integer, GradG As Integer

     Call SetupColors(StartColor, EndColor)
     DoEvents

     picBox.AutoRedraw = True
     picBox.ScaleMode = vbPixels

     If GradType = GradientHorizontal Then Size = picBox.ScaleWidth
     If GradType = GradientVertical Then Size = picBox.ScaleHeight

     For i = 0 To Size
          GradR = ((R2 - R1) / Size * i) + R1
          GradG = ((G2 - G1) / Size * i) + G1
          GradB = ((B2 - B1) / Size * i) + B1

          Color = RGB(GradR, GradG, GradB)

          If GradType = GradientHorizontal Then
               picBox.Line (i, 0)-(i, picBox.ScaleHeight), Color, BF
          ElseIf GradType = GradientVertical Then
               picBox.Line (0, i)-(picBox.ScaleWidth, i), Color, BF
          End If
     Next i

     'picBox.ScaleMode = vbTwips
End Sub

Private Sub SetupColors(ByVal StartColor, EndColor)
     ExtractRGBValues StartColor
     B1 = rBlue
     G1 = rGreen
     R1 = rRed

     ExtractRGBValues EndColor
     B2 = rBlue
     G2 = rGreen
     R2 = rRed
End Sub

Private Function ConvertRGBFormat(ByVal Color As OLE_COLOR) As Long
     TranslateColor Color, 0, ConvertRGBFormat
End Function

Private Function ExtractRGBValues(ByVal vColor As Long)
     rRed = (vColor And &HFF&)
     rGreen = (vColor And &HFF00&) / &H100
     rBlue = (vColor And &HFF0000) / &H10000
End Function

Function hwnd() As Long
     hwnd = UserControl.hwnd
End Function
