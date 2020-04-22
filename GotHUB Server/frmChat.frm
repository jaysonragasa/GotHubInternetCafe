VERSION 5.00
Begin VB.Form frmChat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send Message"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7110
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
   ScaleHeight     =   2040
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send Message"
      Height          =   345
      Left            =   75
      TabIndex        =   1
      Top             =   1620
      Width           =   1215
   End
   Begin VB.TextBox txMsg 
      Height          =   1455
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   90
      Width           =   6945
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSend_Click()
     If frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State <> sckConnected Then
          MsgBox "workstation not connected", vbExclamation
          
          Exit Sub
     End If
     
     frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_CHATMSG & "|" & txMsg.Text
     DoEvents
End Sub
