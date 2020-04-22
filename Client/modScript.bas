Attribute VB_Name = "modScript"
Option Explicit

Public FSo                    As Object
Public WSo                    As Object
Public dTmp                   As String
Public tempScrFyl             As String

Function tX(ByVal pX As Single) As Long
     tX = pX * Screen.TwipsPerPixelX
End Function
Function tY(ByVal pY As Single) As Long
     tY = pY * Screen.TwipsPerPixelY
End Function

Sub InitScript()
     Set FSo = CreateObject("Scripting.FileSystemObject")
     Set WSo = CreateObject("WScript.Shell")
     
     ' get Temp folder
     dTmp = FSo.getspecialfolder(2)
     
     tempScrFyl = FSo.buildpath("C:\", "screen.jpg")
End Sub

