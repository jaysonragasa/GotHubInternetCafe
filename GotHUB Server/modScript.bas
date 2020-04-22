Attribute VB_Name = "modScript"
Option Explicit

Public WSo               As Object
Public FSo               As Object
Public dTmp              As String
Public tempScrFyl        As String

Sub InitScript()
     Set WSo = CreateObject("WScript.Shell")
     Set FSo = CreateObject("Scripting.FileSystemObject")
     dTmp = FSo.getspecialfolder(2)
     
     tempScrFyl = FSo.buildpath("C:\", "screen.jpg")
End Sub
