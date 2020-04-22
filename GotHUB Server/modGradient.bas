Attribute VB_Name = "modGradient"
Option Explicit

Public Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

'The Perfect Gradientizer by Tanner Helland
Sub Gradient(ByRef objPictureBox As Object, _
             ByVal l_Color1 As Long, _
             ByVal l_Color2 As Long, _
             ByVal s_Orientation As String) ' Orientaion: "Vertical" or "Horizontal"
             
     s_Orientation = UCase$(s_Orientation)
             
     'calculation variables for r,g,b gradiency
     Dim VR, VG, VB As Single
     'colors of the picture boxes
     Dim Color1, Color2 As Long
     'r,g,b variables for each picture box
     Dim R, G, B, R2, G2, B2 As Integer
     'calculation variable for extracting the rgb values
     Dim temp As Long
     
     'Tanner Helland forgot the declaration of X and Y
     Dim y As Integer, x As Integer
     
     ' i added some stuff here - JaysonR
     ' ---------------------------------
     On Error Resume Next
     If Not objPictureBox.AutoRedraw Then _
          objPictureBox.AutoRedraw = True
     
     ' translate system color to RGB Color
     l_Color1 = RealColor(l_Color1)
     l_Color2 = RealColor(l_Color2)
     ' ---------------------------------

'     guys we dont need this so i just commented it out - By: JaysonR.
'     ----------------------------------------------
'     Reverses gradient for different mouse buttons
'     If Button = 1 Then
'          Color1 = Picture1.BackColor
'          Color2 = Picture2.BackColor
'     Else
'          Color1 = Picture2.BackColor
'          Color2 = Picture1.BackColor
'     End If
'     ----------------------------------------------

     'extract the r,g,b values from the first picture box
     temp = (l_Color1 And 255)
     R = temp And 255
     temp = Int(l_Color1 / 256)
     G = temp And 255
     temp = Int(l_Color1 / 65536)
     B = temp And 255
     temp = (l_Color2 And 255)
     R2 = temp And 255
     temp = Int(l_Color2 / 256)
     G2 = temp And 255
     temp = Int(l_Color2 / 65536)
     B2 = temp And 255

     If s_Orientation = "VERTICAL" Then
          'create a calculation variable for determining the step between
          'each level of the gradient; this also allows the user to create
          'a perfect gradient regardless of the form size
          'Modfication -------------------------------------------------- by me again.. JaysonR.
          'we are not going to use the form to gradientize,
          'the picture box instead.                    'Orginal
          VR = Abs(R - R2) / objPictureBox.ScaleHeight 'Form1.ScaleHeight
          VG = Abs(G - G2) / objPictureBox.ScaleHeight 'Form1.ScaleHeight
          VB = Abs(B - B2) / objPictureBox.ScaleHeight 'Form1.ScaleHeight
          
          'if the second value is lower then the first value, make the step
          'negative
          If R2 < R Then VR = -VR
          If G2 < G Then VG = -VG
          If B2 < B Then VB = -VB
          
          'run a loop through the form height, incrementing the gradient color
          'according to the height of the line being drawn
                                                 'Orginal
          For y = 0 To objPictureBox.ScaleHeight 'Form1.ScaleHeight
               R2 = R + VR * y
               G2 = G + VG * y
               B2 = B + VB * y
               
               'draw the line and continue through the loop
               'Form1.Line (0, Y)-(Form1.ScaleWidth, Y), RGB(R2, G2, B2) 'Original
               objPictureBox.Line (0, y)-(objPictureBox.ScaleWidth, y), RGB(R2, G2, B2)
          Next y
     Else
          'create a calculation variable for determining the step between
          'each level of the gradient; this also allows the user to create
          'a perfect gradient regardless of the form size
          'Modfication -------------------------------------------------- by me again.. JaysonR.
          'we are not going to use the form to gradientize,
          'the picture box instead.                   'Orginal
          VR = Abs(R - R2) / objPictureBox.ScaleWidth 'Form1.ScaleWidth
          VG = Abs(G - G2) / objPictureBox.ScaleWidth 'Form1.ScaleWidth
          VB = Abs(B - B2) / objPictureBox.ScaleWidth 'Form1.ScaleWidth
          
          'if the second value is lower then the first value, make the step
          'negative
          If R2 < R Then VR = -VR
          If G2 < G Then VG = -VG
          If B2 < B Then VB = -VB
          
          'run a loop through the form height, incrementing the gradient color
          'according to the height of the line being drawn
                                                'Orginal
          For x = 0 To objPictureBox.ScaleWidth 'Form1.ScaleWidth
               R2 = R + VR * x
               G2 = G + VG * x
               B2 = B + VB * x
               
               'draw the line and continue through the loop
               'Form1.Line (X, 0)-(X, Form1.ScaleHeight), RGB(R2, G2, B2) 'Orginal
               objPictureBox.Line (x, 0)-(x, objPictureBox.ScaleHeight), RGB(R2, G2, B2)
          Next x
     End If
End Sub

Public Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     
     Col = TranslateColor(Color, 0, RealColor)
End Function

