VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   75
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   75
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mMenu 
      Caption         =   "menu"
      Begin VB.Menu smMenu 
         Caption         =   "Connection Settings"
         Index           =   0
      End
      Begin VB.Menu smMenu 
         Caption         =   "Change Theme"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu smMenu 
         Caption         =   "Disconnect"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub smMenu_Click(Index As Integer)
     If Index = 0 Then        ' connection settings
          frmSettings.SSTab1.Tab = 0
          frmSettings.Show vbModal, Me
          
     ElseIf Index = 1 Then    ' change theme
          frmSettings.SSTab1.Tab = 1
          frmSettings.Show vbModal, Me
          
     ElseIf Index = 2 Then
          frmMain.DisconnectToServer
          End
     End If
     
     
End Sub
