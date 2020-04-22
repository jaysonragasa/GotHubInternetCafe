VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{E6E03A98-C7DC-4FCE-800D-724A332410A9}#1.0#0"; "LaVolpeButtons.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "GotHub? - Client Application"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picINetStat 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   1740
      ScaleHeight     =   5340
      ScaleWidth      =   7665
      TabIndex        =   5
      Top             =   1755
      Visible         =   0   'False
      Width           =   7665
      Begin VB.TextBox txName 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   480
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "username"
         Top             =   390
         Width           =   4335
      End
      Begin LaVolpeButtons.lvButtons_H btnRqLO 
         Height          =   360
         Left            =   2985
         TabIndex        =   15
         Top             =   4875
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   635
         Caption         =   "&Request Logout"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":57E2
         cBack           =   -2147483633
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   2985
         Picture         =   "frmMain.frx":5B7C
         Top             =   4470
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please click the 'Request Logout' button."
         Height          =   210
         Index           =   1
         Left            =   3315
         TabIndex        =   17
         Top             =   4590
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If your done using this workstation"
         Height          =   210
         Index           =   0
         Left            =   3315
         TabIndex        =   16
         Top             =   4410
         Width           =   2505
      End
      Begin VB.Label lblINRent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Rental: 0.00 Php"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2985
         TabIndex        =   14
         Top             =   3270
         Width           =   2340
      End
      Begin VB.Image Image4 
         Height          =   30
         Left            =   3015
         Picture         =   "frmMain.frx":5F06
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.000 Php"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   2985
         TabIndex        =   13
         Top             =   3855
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2985
         TabIndex        =   12
         Top             =   3615
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00h 00m"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   2985
         TabIndex        =   11
         Top             =   2700
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Used"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2985
         TabIndex        =   10
         Top             =   2415
         Width           =   1005
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00 am/pm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   2985
         TabIndex        =   9
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log In Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2985
         TabIndex        =   8
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Image Image3 
         Height          =   1935
         Left            =   2790
         Picture         =   "frmMain.frx":63A0
         Top             =   1140
         Width           =   4305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Workstation User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   435
         Left            =   2775
         TabIndex        =   6
         Top             =   15
         Width           =   4365
      End
      Begin VB.Image Image1 
         Height          =   3480
         Left            =   15
         Picture         =   "frmMain.frx":6AB4
         Top             =   405
         Width           =   3780
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10590
      TabIndex        =   3
      Top             =   7320
      Width           =   10590
      Begin VB.PictureBox picBtns 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   4635
         ScaleHeight     =   555
         ScaleWidth      =   5505
         TabIndex        =   18
         Top             =   0
         Width           =   5505
         Begin VB.TextBox txPass 
            BackColor       =   &H80000018&
            ForeColor       =   &H80000017&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   30
            PasswordChar    =   "*"
            TabIndex        =   23
            Top             =   210
            Width           =   1410
         End
         Begin LaVolpeButtons.lvButtons_H btns 
            Height          =   495
            Index           =   0
            Left            =   1500
            TabIndex        =   19
            Top             =   30
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Disconnect"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   -2147483639
            cFHover         =   -2147483639
            cBhover         =   -2147483635
            Focus           =   0   'False
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmMain.frx":AC77
            ImgSize         =   24
            cBack           =   -2147483646
         End
         Begin LaVolpeButtons.lvButtons_H btns 
            Height          =   495
            Index           =   1
            Left            =   2970
            TabIndex        =   20
            Top             =   30
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            Caption         =   "&Settings"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   -2147483639
            cFHover         =   -2147483639
            cBhover         =   -2147483635
            Focus           =   0   'False
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmMain.frx":BD89
            ImgSize         =   24
            cBack           =   -2147483646
         End
         Begin LaVolpeButtons.lvButtons_H btns 
            Height          =   495
            Index           =   2
            Left            =   4200
            TabIndex        =   21
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   873
            Caption         =   "E&xit Client"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   -2147483639
            cFHover         =   -2147483639
            cBhover         =   -2147483635
            Focus           =   0   'False
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmMain.frx":C483
            ImgSize         =   24
            cBack           =   -2147483646
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter password"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   45
            TabIndex        =   22
            Top             =   0
            Width           =   1170
         End
      End
      Begin VB.Image Icons 
         Height          =   240
         Index           =   6
         Left            =   780
         Picture         =   "frmMain.frx":CB7D
         Top             =   1515
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Icons 
         Height          =   240
         Index           =   5
         Left            =   465
         Picture         =   "frmMain.frx":D107
         Top             =   1530
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Icons 
         Height          =   240
         Index           =   4
         Left            =   165
         Picture         =   "frmMain.frx":D491
         Top             =   1545
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Icons 
         Height          =   240
         Index           =   3
         Left            =   2055
         Picture         =   "frmMain.frx":DA1B
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Icons 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":DFA5
         Top             =   1140
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Icons 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmMain.frx":E0EF
         Top             =   1155
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Icons 
         Height          =   240
         Index           =   0
         Left            =   165
         Picture         =   "frmMain.frx":E239
         Top             =   1170
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Ico 
         Height          =   240
         Left            =   75
         Picture         =   "frmMain.frx":E7C3
         Top             =   165
         Width           =   240
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Idle..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   390
         TabIndex        =   4
         Top             =   105
         Width           =   795
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      Picture         =   "frmMain.frx":ED4D
      ScaleHeight     =   2655
      ScaleWidth      =   7965
      TabIndex        =   2
      Top             =   4650
      Width           =   7965
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   3300
      Top             =   2400
   End
   Begin VB.PictureBox picInfo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5220
      Left            =   105
      Picture         =   "frmMain.frx":46913
      ScaleHeight     =   5220
      ScaleWidth      =   6045
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   6045
      Begin LaVolpeButtons.lvButtons_H btnLI 
         Height          =   405
         Left            =   465
         TabIndex        =   1
         Top             =   3900
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   714
         Caption         =   "Start Now"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":AD8E5
         cBack           =   -2147483633
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   105
         Picture         =   "frmMain.frx":ADC7F
         Top             =   1320
         Width           =   360
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   5505
      Top             =   555
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   5535
      Top             =   1035
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5055
      Top             =   525
   End
   Begin VB.Image imgHeader 
      Height          =   2520
      Left            =   150
      Picture         =   "frmMain.frx":AE3E9
      Top             =   345
      Visible         =   0   'False
      Width           =   15360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim ConTries                  As Long
Dim i                         As Long
Dim Reconnect                 As Boolean

Dim oX                        As Single
Dim oY                        As Single

Private Sub btnLI_Click()
     sckClient.SendData CONN_LOGEDIN: DoEvents
End Sub

Private Sub btnRqLO_Click()
     sckClient.SendData CONN_LOGOUT: DoEvents
End Sub

Private Sub btns_Click(Index As Integer)
     If Setting.SettingsPassword = txPass Then
          If Index = 0 Then
               DisconnectToServer
          ElseIf Index = 1 Then
               frmSettings.SSTab1.Tab = 0
               frmSettings.Show vbModal, Me
          ElseIf Index = 2 Then
               'Call DisableTaskbar(False)
     
               Unload Me
          End If
     Else
          MsgBox "Either you did not enter your settings password or password did not match" + vbCrLf + _
                 "Please enter your password on the textfield on the left side of Disconnect Button.", vbExclamation
     End If
End Sub

' MESSAGE PATTERN: <Constant Integer Value or Message will send on the server>|<any string data>|<password>
'                  we will be supplying password for every data sended. this will serve as additional
'                  security on the server

Private Sub Form_Load()
     Dim clsDesk         As New DesktopArea
     
     BackColor = vbWhite
     
     clsDesk.PositionForm Me, H_FULL, V_FULL
     
     picMain.Move (Width - picMain.Width) / 2, (Height - picMain.Height) / 2
     picInfo.Move 0, 0
     imgHeader.Move 0, 0
     picINetStat.Move (Width - picINetStat.Width), imgHeader.Height - tX(20)
     picBtns.Move Width - picBtns.Width
     
     Call InitScript
     Call InitializeSettings
     
     'Call PrepareThemeSupport
     
     'Call FixThemeSupport(Controls)
     
     ' right-click on the procedure name and click the "Definition" menu for more details.
     'Call StayBottom(hwnd)
     
     ' right-click on the procedure name and click the "Definition" menu for more details.
     'Call ConnectToServer
     
     'Call DisableTaskbar(False)
     
     Call StylePasswordField(txPass)
     
     i = 4
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     sckClient.Close
     
     KillTimer hwnd, 0
End Sub

' monitor connection status
Private Sub Timer1_Timer()
     Dim pMin            As Double
     Dim cTime           As String
     Dim TtlInetAmnt     As Currency
     Dim TimeUsed        As String
     Dim tParse()        As String
     
     If sckClient.State <> sckConnected Then
          If i = 7 Then i = 4
          Ico.Picture = Icons(i)
          i = i + 1
     End If

     ' i dont no why, if you terminate the Server Application,
     ' the client state will be on sckClosing State (Constant Value = 8)
     ' may be because; for us, to manualy close the connection??

     ' so if the client state is sckClosing State then
     If sckClient.State = sckClosing Then
          ' close the client connection
          sckClient.Close
          ' shout the statement sckClient.Close to do it nicely
          DoEvents

          lblStat.Caption = "Disconnected from Server..."
          
          ' reconnect again
          Call ConnectToServer
     End If
     
     If LogedIn Then
          TimeUsed = Format(TimeValue(Time) - TimeValue(lblInfo(0).Caption), "hh:mm:ss")
          
          tParse = Split(TimeUsed, ":")
          
          lblInfo(1).Caption = tParse(0) & "h, " & _
                               tParse(1) & "m, " & _
                               tParse(2) & "s"
          
          pMin = CCur(lblINRent.Tag) / 60
          
          cTime = lblInfo(0).Caption
          cTime = DateDiff("s", cTime, Time)
          cTime = cTime \ 60
          
          TtlInetAmnt = cTime * pMin
     
          lblInfo(2).Caption = FormatNumber(CCur(TtlInetAmnt), 2) & " Php"
     End If
     
     'Call StayBottom(hwnd)
End Sub

Private Sub Timer2_Timer()
     If sckClient.State <> sckConnected Then
          Reconnect = True
     
          Call sckClient_Error(0, vbNullString, 0, vbNullString, vbNullString, 0, False)
     End If
End Sub


Private Sub sckClient_Connect()
     lblStat.Caption = "waiting for authentication..."
     
     Ico.Picture = Icons(1).Picture
     
     ConTries = 0
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
     Dim sData           As String
     Dim msg()           As String
     Dim Chunks          As String
     
     sData = vbNullString
     sckClient.GetData sData
     msg = Split(sData, "|")
     
     If sData = CONN_CONNECTED Then
          lblStat.Caption = "Connected"
          Ico.Picture = Icons(3).Picture
          
          sckClient.SendData CONN_REQUESTSTATUS
          DoEvents
          
'          Call LockMe(hwnd, True)
               
     ElseIf sData = CONN_DISCONNECTED Then
          sckClient.Close
               
          lblStat.Caption = "Disconnected"
               
     ElseIf msg(0) = CONN_LOGIN Then
          'LogedIn = True
          'lblStat.Caption = "LogIn Time: " & msg(1) & ", Time Used: 00h 00m"
          'Ico.Picture = Icons(2).Picture
          
          imgHeader.Visible = False
          picMain.Visible = False
          picInfo.Visible = True
               
     ElseIf sData = CONN_LOGOUT Then
          Ico.Picture = Icons(3).Picture
          lblStat.Caption = "Connected"
          txName.Text = vbNullString
          
          LogedIn = False
          
          picMain.Visible = True
          imgHeader.Visible = False
          picInfo.Visible = False
          picINetStat.Visible = False
          
     ElseIf sData = CONN_CANCEL Then
          sckClient.Close
               
          lblStat.Caption = "Canceled": DoEvents
          lblStat.Caption = "Connected": DoEvents
          txName.Text = vbNullString
          Ico.Picture = Icons(3).Picture
          
          LogedIn = False
          
          picMain.Visible = True
          imgHeader.Visible = False
          picInfo.Visible = False
          picINetStat.Visible = False
               
     ElseIf msg(0) = "UNAME" Then
          LogedIn = True
          
          txName.Text = msg(1)
          
          picInfo.Visible = False
          picMain.Visible = False
          
          imgHeader.Visible = True
          picINetStat.Visible = True
          
          lblInfo(0).Caption = Time
          lblINRent.Caption = "Internet Rental: " & msg(2) & " Php"
          lblINRent.Tag = msg(2)
               
     ElseIf msg(0) = "TIMER" Then
          lblInfo(0).Caption = msg(1)
          lblInfo(1).Caption = msg(2)
          
     ElseIf msg(0) = CONN_CHATMSG Then
          frmChat.txMsg.Text = msg(1)
          
          If Not frmChat.Visible Then frmChat.Show vbModal, Me
          
     ElseIf sData = CONN_ENUMWIN Then
          sEnumed = vbNullString
          
          EnumWindows AddressOf EnumWnd, 0
          DoEvents
          
          sEnumed = Left$(sEnumed, Len(sEnumed) - 1)
          
          sckClient.SendData CONN_SENDENUMWIN + "=" + sEnumed
     ElseIf msg(0) = CONN_CLOSEAPP Then
          Call CloseApplication(CLng(msg(1)))
     
     ElseIf msg(0) = CONN_REQUESTSTATUS Then
          LogedIn = True
          
          txName.Text = msg(1)
          
          lblInfo(0).Caption = msg(2)
          lblInfo(1).Caption = msg(3)
          lblINRent.Caption = "Internet Rental: " & msg(4) & " Php"
          lblINRent.Tag = msg(4)
          
          picInfo.Visible = False
          picMain.Visible = False
          
          imgHeader.Visible = True
          picINetStat.Visible = True
          
     ElseIf sData = CONN_CAPTURESCREEN Then
          With frmTemp
               .temp.Cls
               .temp.Picture = CaptureScreen()
               .temp.Refresh
          
               .picRes.Cls
          
               .picRes.PaintPicture .temp, 0, 0, .picRes.Width, .picRes.Height
               
               Set .Image1.Picture = ConvertToPicture(.picRes.Image)
               
               SavePicture .Image1.Picture, tempScrFyl
               DoEvents
               
               sckClient.SendData CONN_FILESTAT & "=" & "1" & "=" & FileLen(tempScrFyl)
               DoEvents
               
               Open tempScrFyl For Binary As #1
               DoEvents
               
               Sleep 200
          End With
     
     ElseIf msg(0) = CONN_FILESTAT Then
          If msg(1) = "2" Then
               Do While Not EOF(1)
                    If sckClient.State <> sckConnected Then Exit Sub
                    
                    Chunks$ = Input(MAX_CHUNK, #1)
                    DoEvents
               
                    sckClient.SendData CONN_FILESTAT & "=" & "3" & "=" & Chunks$
                    DoEvents
                    
                    Sleep 200
               Loop
               
               Sleep 1000
               
               Close #1
               DoEvents
               
               sckClient.SendData CONN_FILESTAT & "=" & "4"
               DoEvents
               
               Kill tempScrFyl
               DoEvents
          End If
     
     ElseIf msg(0) = CONN_LOCKWS Then
          LockMe Me.hwnd, CBool(msg(1))
     ElseIf msg(0) = CONN_MONITOROFF Then
          MonitorOff Me.hwnd, CBool(msg(1))
     ElseIf msg(0) = CONN_SHUTDOWN Then
          ShutdownBy CLng(msg(1))
     ElseIf msg(0) = CONN_MOUSECLICK Then
          ClickThis CLng(msg(1)), CLng(msg(2))
          
     End If
End Sub

' event used when connection has an error (e.g. cannot connect to server)
Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
     lblStat.Caption = "Server not found": DoEvents
     
     ' close the client connection
     sckClient.Close
     
     ' wait for 3 seconds and
     If Reconnect Then
          ' then reconnect
          Call ConnectToServer
          
          Reconnect = False
     End If
End Sub

Sub ConnectToServer()
     
     ' count, how many times the connection was tried
     ConTries = ConTries + 1
     
     lblStat.Caption = "Connection Tries [ " & ConTries & " ] Connecting..."
     
     sckClient.Close
     
     ' connect to server <sckClient.Connect <IP Address>, <Port>
     sckClient.Connect Setting.Server_IPAddress, Setting.Server_Port
                       
     ' shout the statement sckClient.Connect to do it nicely
     DoEvents
End Sub

Sub DisconnectToServer()
     If sckClient.State = sckConnected Then
          sckClient.SendData CStr(CONN_DISCONNECT)
          
          lblStat.Caption = "Disconnecting To Server..."
          
          DoEvents
     End If
End Sub
