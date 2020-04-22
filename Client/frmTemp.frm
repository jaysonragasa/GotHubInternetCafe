VERSION 5.00
Begin VB.Form frmTemp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox temp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   135
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   3600
      Width           =   1125
   End
   Begin VB.PictureBox picRes 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   1440
      ScaleHeight     =   219
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   0
      Top             =   3630
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   3285
      Left            =   225
      Top             =   195
      Width           =   4320
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

