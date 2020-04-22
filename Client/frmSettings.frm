VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client Settings"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   2460
      Left            =   540
      ScaleHeight     =   2400
      ScaleWidth      =   6015
      TabIndex        =   5
      Top             =   2565
      Width           =   6075
      Begin VB.TextBox txSecPass 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1515
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   480
         Width           =   2550
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Go"
         Height          =   195
         Left            =   4380
         TabIndex        =   13
         Top             =   525
         Width           =   195
      End
      Begin VB.Image imgGo 
         Height          =   240
         Left            =   4110
         Picture         =   "frmSettings.frx":038A
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   7
         Top             =   510
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter the password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   690
         TabIndex        =   6
         Top             =   90
         Width           =   3300
      End
      Begin VB.Image Image1 
         Height          =   795
         Index           =   1
         Left            =   165
         Picture         =   "frmSettings.frx":0714
         Top             =   135
         Width           =   705
      End
   End
   Begin VB.CommandButton btns 
      Caption         =   "&OK"
      Height          =   345
      Index           =   1
      Left            =   5010
      TabIndex        =   4
      Top             =   5715
      Width           =   1215
   End
   Begin VB.CommandButton btns 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   0
      Left            =   6270
      TabIndex        =   3
      Top             =   5715
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Main Settings"
      TabPicture(0)   =   "frmSettings.frx":2526
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txSettings(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txSettings(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "mskEdt"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txSettings(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Theme"
      TabPicture(1)   =   "frmSettings.frx":2542
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.TextBox txSettings 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   5400
         TabIndex        =   15
         Tag             =   "WorkstationLtr"
         Top             =   540
         Width           =   915
      End
      Begin MSMask.MaskEdBox mskEdt 
         Height          =   285
         Left            =   1620
         TabIndex        =   12
         Top             =   540
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   15
         Format          =   "###.###.###.###"
         Mask            =   "###.###.###.###"
         PromptChar      =   "0"
      End
      Begin VB.TextBox txSettings 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   10
         Tag             =   "SettingsPassword"
         Top             =   930
         Width           =   2205
      End
      Begin VB.TextBox txSettings 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   2
         Tag             =   "ServerIPAddress"
         Top             =   930
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Workstation Letter:"
         Height          =   195
         Index           =   2
         Left            =   3945
         TabIndex        =   14
         Top             =   570
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4155
         MouseIcon       =   "frmSettings.frx":255E
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   960
         Width           =   1125
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   3870
         Picture         =   "frmSettings.frx":3228
         Top             =   945
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Settings Password:"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   9
         Top             =   975
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP Address:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   570
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btns_Click(Index As Integer)
     Dim Cntrls          As Control
     
     If Index = 1 Then
          If IsValid Then
               For Each Cntrls In Controls
                    If TypeOf Cntrls Is Textbox Then
                         Call WriteData(Cntrls.Tag, Cntrls.Text)
                    End If
               Next
               
               With Setting
                    .Server_IPAddress = txSettings(0).Text
                    .SettingsPassword = txSettings(1).Text
                    .server_WorkstationLtr = txSettings(2).Text
               End With
               
               Unload Me
          End If
     ElseIf Index = 0 Then
          Unload Me
     End If
End Sub

Function IsValid() As Boolean
     Dim Cntrls          As Control
     Dim tb_cnt          As Integer
     Dim cnt             As Integer
     
     For Each Cntrls In Controls
          If TypeOf Cntrls Is Textbox Then
               tb_cnt = tb_cnt + 1
               
               If Cntrls.Text <> vbNullString Then
                    cnt = cnt + 1
               End If
          End If
     Next
     
     If cnt = tb_cnt Then
          IsValid = True
     Else
          IsValid = False
     End If
End Function

Private Sub Form_Activate()
     txSecPass.SetFocus
End Sub

Private Sub Form_Load()
     Dim tmp()      As String
     Dim IP         As String
     
     tmp = Split(Setting.Server_IPAddress, ".")
     
     IP = Right$("000", 3 - Len(tmp(0))) & tmp(0) & "."
     IP = IP + Right$("000", 3 - Len(tmp(1))) & tmp(1) & "."
     IP = IP + Right$("000", 3 - Len(tmp(2))) & tmp(2) & "."
     IP = IP + Right$("000", 3 - Len(tmp(2))) & tmp(3)
     
     mskEdt.Text = IP
     txSettings(1).Text = Setting.SettingsPassword
     txSettings(2).Text = Setting.server_WorkstationLtr
     
     
     Picture1.Move 0, 0, ScaleWidth, ScaleHeight
     Picture1.ZOrder vbBringToFront
     
     Call StylePasswordField(txSettings(1))
     Call StylePasswordField(txSecPass)
End Sub

Private Sub imgGo_Click()
     If txSecPass.Text = Setting.SettingsPassword Then
          Picture1.Visible = False
     Else
          MsgBox "Invald Password", vbCritical
     End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     txSettings(1).PasswordChar = vbNullString
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     txSettings(1).PasswordChar = "*"
End Sub

Private Sub mskEdt_Change()
     txSettings(0).Text = mskEdt.Text
End Sub

Private Sub txSecPass_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = Asc(vbCrLf) Then
          Call imgGo_Click
     End If
End Sub
