VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMonitor 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monitor Workstation"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   5595
      Index           =   1
      Left            =   2775
      ScaleHeight     =   5595
      ScaleWidth      =   5895
      TabIndex        =   14
      Top             =   705
      Width           =   5895
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   225
         Left            =   1950
         TabIndex        =   15
         Top             =   5145
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click the screen to control the workstation."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1095
         TabIndex        =   25
         Top             =   405
         Width           =   3645
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refreshing... 0%"
         Height          =   195
         Left            =   630
         TabIndex        =   16
         Top             =   5160
         Width           =   1260
      End
      Begin VB.Image imgScr 
         Height          =   3285
         Left            =   735
         Picture         =   "frmMonitor.frx":038A
         Top             =   720
         Width           =   4320
      End
      Begin VB.Image Image1 
         Height          =   4950
         Left            =   285
         Picture         =   "frmMonitor.frx":83C6
         Stretch         =   -1  'True
         Top             =   135
         Width           =   5325
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   5955
      Left            =   0
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   172
      TabIndex        =   4
      Top             =   570
      Width           =   2580
      Begin GoHUB_Server.InfoHeader ctl_IH 
         Height          =   1980
         Index           =   2
         Left            =   60
         TabIndex        =   18
         Top             =   3585
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   3493
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   -2147483643
         FrameBorderColor=   10070188
         GradientStyle   =   0
         LeftColor       =   14898176
         RightColor      =   14215660
         Margin          =   4
         BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         Caption         =   "Other Tools"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMonitor.frx":E240
         EdgeRadiusSize  =   8
         RoundEdgeStyle  =   6
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000005&
            Caption         =   "Force Shutdown"
            Height          =   195
            Index           =   3
            Left            =   390
            TabIndex        =   24
            Tag             =   "4"
            Top             =   1365
            Width           =   1890
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000005&
            Caption         =   "Reboot"
            Height          =   195
            Index           =   2
            Left            =   390
            TabIndex        =   23
            Tag             =   "2"
            Top             =   1140
            Width           =   1890
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000005&
            Caption         =   "Shutdown"
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   22
            Tag             =   "1"
            Top             =   915
            Width           =   1890
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000005&
            Caption         =   "Log Off"
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   21
            Tag             =   "0"
            Top             =   675
            Width           =   1890
         End
         Begin GoHUB_Server.AdvanceLabelControl ctlLbl 
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   1620
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   476
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Shutdown"
            NormalIcon      =   "frmMonitor.frx":E7DA
            HasIcon         =   -1  'True
            MouseIcon       =   "frmMonitor.frx":EB74
            AutoResize      =   -1  'True
            MouseIcon       =   "frmMonitor.frx":F84E
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "System"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   390
            Width           =   645
         End
      End
      Begin GoHUB_Server.InfoHeader ctl_IH 
         Height          =   2190
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   1290
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   3863
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   -2147483643
         FrameBorderColor=   10070188
         GradientStyle   =   0
         LeftColor       =   14898176
         RightColor      =   14215660
         Margin          =   4
         BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         Caption         =   "Screen Monitor"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMonitor.frx":10528
         EdgeRadiusSize  =   8
         RoundEdgeStyle  =   6
         Begin GoHUB_Server.AdvanceLabelControl ctlLbl 
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Tag             =   "0"
            Top             =   1830
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   476
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Monitor Off"
            NormalIcon      =   "frmMonitor.frx":108C2
            HasIcon         =   -1  'True
            MouseIcon       =   "frmMonitor.frx":10C5C
            AutoResize      =   -1  'True
            MouseIcon       =   "frmMonitor.frx":11936
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000005&
            Caption         =   "Entire Screen"
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   13
            Top             =   1245
            Width           =   1875
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000005&
            Caption         =   "Form"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   12
            Top             =   1005
            Width           =   1875
         End
         Begin GoHUB_Server.AdvanceLabelControl ctlLbl 
            Height          =   270
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   1530
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   476
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Refresh Screen"
            NormalIcon      =   "frmMonitor.frx":12610
            HasIcon         =   -1  'True
            MouseIcon       =   "frmMonitor.frx":12BAA
            AutoResize      =   -1  'True
            MouseIcon       =   "frmMonitor.frx":13884
         End
         Begin GoHUB_Server.AdvanceLabelControl ctlLbl 
            Height          =   270
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Tag             =   "0"
            Top             =   405
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   476
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Lock Workstation"
            NormalIcon      =   "frmMonitor.frx":1455E
            HasIcon         =   -1  'True
            MouseIcon       =   "frmMonitor.frx":14AF8
            AutoResize      =   -1  'True
            MouseIcon       =   "frmMonitor.frx":157D2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Screen Capture"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   735
            Width           =   1875
         End
      End
      Begin GoHUB_Server.InfoHeader ctl_IH 
         Height          =   1110
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   75
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1958
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   -2147483643
         FrameBorderColor=   10070188
         GradientStyle   =   0
         LeftColor       =   14898176
         RightColor      =   14215660
         Margin          =   4
         BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         Caption         =   "Windows Enumeration"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMonitor.frx":164AC
         EdgeRadiusSize  =   8
         RoundEdgeStyle  =   6
         Begin GoHUB_Server.AdvanceLabelControl ctlLbl 
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   476
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Close Application"
            NormalIcon      =   "frmMonitor.frx":16A46
            HasIcon         =   -1  'True
            MouseIcon       =   "frmMonitor.frx":16FE0
            AutoResize      =   -1  'True
            MouseIcon       =   "frmMonitor.frx":17CBA
         End
         Begin GoHUB_Server.AdvanceLabelControl ctlLbl 
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   476
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Refresh List"
            NormalIcon      =   "frmMonitor.frx":18994
            HasIcon         =   -1  'True
            MouseIcon       =   "frmMonitor.frx":18F2E
            AutoResize      =   -1  'True
            MouseIcon       =   "frmMonitor.frx":19C08
         End
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   5595
      Index           =   0
      Left            =   2790
      ScaleHeight     =   5595
      ScaleWidth      =   5895
      TabIndex        =   2
      Top             =   780
      Width           =   5895
      Begin MSComctlLib.ListView lvEnumList 
         Height          =   5460
         Left            =   60
         TabIndex        =   3
         Top             =   75
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   9631
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Window Title"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox ico 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3810
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   6225
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7785
      Top             =   5415
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList TB_IL 
      Left            =   7890
      Top             =   645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitor.frx":1A8E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitor.frx":1AE7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitor.frx":1B416
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitor.frx":1B9B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1005
      ButtonWidth     =   2990
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "TB_IL"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Windows Enumeration"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Screen Monitor"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEL_hWnd             As String
Dim SEL_Index            As String
Dim SD_Index             As Long
Dim Refreshed            As Boolean

Private Sub ctlLbl_Click(Index As Integer)
     If frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State <> sckConnected Then
          MsgBox "workstation not connected.", vbExclamation
          
          Exit Sub
     End If
     
     If Index = 0 Then        ' refresh list
          frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_ENUMWIN
          DoEvents
          
     ElseIf Index = 1 Then    ' close app
          frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_CLOSEAPP & "|" & SEL_hWnd
          DoEvents
          
          Call ctlLbl_Click(0)
          DoEvents
          
     ElseIf Index = 2 Then    ' lock screen
          If ctlLbl(Index).Tag = 0 Then
               frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_LOCKWS & "|" & "1"
               DoEvents
               
               ctlLbl(Index).Caption = "Unlock Workstation"
               ctlLbl(Index).Tag = 1
          ElseIf ctlLbl(Index).Tag = 1 Then
               frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_LOCKWS & "|" & "0"
               DoEvents
               
               ctlLbl(Index).Caption = "Lock Workstation"
               ctlLbl(Index).Tag = 0
          End If
          
     ElseIf Index = 3 Then    ' refresh screen
          If Option1(0).Value = True Then
               
               'Refreshed = True
          ElseIf Option1(1).Value = True Then
               frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_CAPTURESCREEN
               DoEvents
               
               Refreshed = True
               
               ctlLbl(Index).Enabled = False
          Else
               MsgBox "please select the screen capture type", vbInformation
          End If
          
     ElseIf Index = 4 Then    ' monitor off
          If ctlLbl(Index).Tag = 0 Then
               frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_MONITOROFF & "|" & "1"
               DoEvents
               
               ctlLbl(Index).Caption = "Monitor On"
               ctlLbl(Index).Tag = 1
          ElseIf ctlLbl(Index).Tag = 1 Then
               frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_MONITOROFF & "|" & "0"
               DoEvents
               
               ctlLbl(Index).Caption = "Monitor Off"
               ctlLbl(Index).Tag = 0
          End If
     
     ElseIf Index = 5 Then    ' shutdown
          frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_SHUTDOWN & "|" & SD_Index
          DoEvents
     ElseIf Index = 6 Then
     End If
End Sub

Private Sub Form_Load()
     Dim i               As Integer
     
     Width = 8565: Height = 7260
     
     ' refresh enumlist
     Call ctlLbl_Click(0)
     Call StyleStaticEdge(lvEnumList)
     
     Call Gradient(Picture2, RGB(122, 161, 230), vbWhite, "VERTICAL")
     
     For i = 0 To ctl_IH.Count - 1
          ctl_IH(i).EdgeRadiusSize = 6
          ctl_IH(i).LeftColor = vbWhite
          ctl_IH(i).RightColor = RGB(199, 211, 247)
          ctl_IH(i).ForeColor = RGB(33, 93, 198)
          ctl_IH(i).FrameBorderColor = RGB(199, 211, 247)
          ctl_IH(i).FrameBackColor = vbWhite
          ctl_IH(i).RoundEdgePosition = REP_Top
          DoEvents
     Next i
     
     For i = Frames.Count - 1 To 0 Step -1
          Frames(i).Move Picture2.Width, Toolbar1.Height
          Frames(i).Height = Height - Frames(i).Top
          Frames(i).BackColor = vbWhite
          Frames(i).ZOrder vbBringToFront
     Next i
     
     Refreshed = False
End Sub

Private Sub imgScr_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim percX      As Long
     Dim percY      As Long
     
     If Not Refreshed Then
          MsgBox "Please refreshed the screen by selecting the Screen Type and click the Refresh Button.", vbInformation
          
          Exit Sub
     End If
     
     If Button = 1 Then
          percX = CInt((x / imgScr.Width) * 100)
          percY = CInt((y / imgScr.Height) * 100)
     
          frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_MOUSECLICK & "|" & percX & "|" & percY
          DoEvents
          
          Sleep 200
     End If
End Sub

Private Sub lvEnumList_ItemClick(ByVal Item As MSComctlLib.ListItem)
     SEL_hWnd = Item.Tag
     SEL_Index = Item.Index
End Sub

Private Sub Option2_Click(Index As Integer)
     SD_Index = Option2(Index).Tag
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     If Button.Index = 1 Then           ' refresh enum list
          Frames(0).ZOrder vbBringToFront
     ElseIf Button.Index = 2 Then       ' refresh screen capture
          Frames(1).ZOrder vbBringToFront
     ElseIf Button.Index = 4 Then       ' close app
     End If
End Sub
