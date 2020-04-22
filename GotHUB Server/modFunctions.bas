Attribute VB_Name = "modFunctions"
Option Explicit

Function tX(ByVal pX As Single) As Long
     tX = pX * Screen.TwipsPerPixelX
End Function
Function tY(ByVal pY As Single) As Long
     tY = pY * Screen.TwipsPerPixelY
End Function

Sub AutoHighlight(ByVal ctlTextBox As TextBox)
     ctlTextBox.SelStart = 0
     ctlTextBox.SelLength = Len(ctlTextBox)
End Sub

Sub HideAllShadowFrame(ByVal srcForm As Form)
     Dim Cntrl      As Control
     
     For Each Cntrl In srcForm
          If TypeOf Cntrl Is PictureBox Then
               If InStr(1, LCase$(Cntrl.Name), "shadow") <> 0 Then
                    Cntrl.Visible = False
               End If
          End If
     Next
End Sub

Sub HideAllFloatingControl(ByVal srcForm As Form)
     Dim Cntrl      As Control
     
     For Each Cntrl In srcForm
          'If TypeOf Cntrl Is InfoHeader Then
               If InStr(1, LCase$(Cntrl.Name), "float") <> 0 Then
                    Cntrl.Visible = False
               End If
          'End If
     Next
End Sub

Sub ShowInformation(ByVal Info As String)
     frmMain.lblInfo.Caption = Info
End Sub

Public Function IsIDE() As Boolean
     On Error GoTo ErrHandler
     
     'because debug statements are ignored when
     'the app is compiled, the next statment will
     'never be executed in the EXE.
     Debug.Print 1 / 0
     
     IsIDE = False
     
     Exit Function
ErrHandler:
     'If we get an error then we are
     'running in IDE / Debug mode
     IsIDE = True
End Function

Sub Pause(ByVal interval As Long)
     Dim l     As Long
     
     For l = 1 To 500000: DoEvents: Next l
End Sub
