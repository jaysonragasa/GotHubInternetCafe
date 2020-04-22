VERSION 5.00
Begin VB.Form frmUserName 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PC Username"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3180
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btns 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   2250
      TabIndex        =   2
      Top             =   720
      Width           =   810
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   315
      Width           =   2940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PC Username"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "frmUserName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btns_Click(Index As Integer)
     If Index = 0 Then
          frmMain.lvClients.ListItems(ItmDetails.ItemIndex).SubItems(1) = Text1.Text
     End If
     
     Unload Me
End Sub

Private Sub Text1_Change()
     If Text1.Text <> vbNullString Then
          btns(0).Enabled = True
     Else: btns(0).Enabled = False
     End If
End Sub
