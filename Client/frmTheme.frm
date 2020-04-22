VERSION 5.00
Begin VB.Form frmDesktop 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesk 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   6285
      Left            =   540
      ScaleHeight     =   6285
      ScaleWidth      =   8820
      TabIndex        =   0
      Top             =   390
      Width           =   8820
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     Dim clsDesk         As New DesktopArea
     
     clsDesk.PositionForm Me, H_FULL, V_FULL
     
     BackColor = vbWhite
     picDesk.Move 0, 0, Width, Height
     
     StayBottom hwnd
End Sub

Private Sub Timer1_Timer()
     StayBottom hwnd
End Sub
