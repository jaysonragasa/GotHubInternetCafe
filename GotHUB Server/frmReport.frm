VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report Options"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3210
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
   ScaleHeight     =   4890
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btns 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   1
      Left            =   675
      TabIndex        =   8
      Top             =   4470
      Width           =   1215
   End
   Begin VB.CommandButton btns 
      Caption         =   "&Show Report"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   1935
      TabIndex        =   7
      Top             =   4470
      Width           =   1215
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   4305
      Index           =   1
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   7594
      InfoHeaderStyle =   1
      HeaderHeight    =   26
      FrameBackColor  =   16777215
      FrameBorderColor=   10070188
      GradientStyle   =   1
      LeftColor       =   14215660
      RightColor      =   16777215
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
      ForeColor       =   -2147483630
      Caption         =   "Options"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   0   'False
      Picture         =   "frmReport.frx":0000
      EdgeRadiusSize  =   0
      RoundEdgeStyle  =   7
      Begin MSComCtl2.MonthView MonthView 
         Height          =   2370
         Left            =   345
         TabIndex        =   6
         Top             =   1800
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   67764225
         TitleBackColor  =   -2147483626
         CurrentDate     =   38422
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Yearly"
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   4
         Top             =   1185
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Monthly"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   3
         Top             =   945
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Daily"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   2
         Top             =   705
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1515
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Report Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   465
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelInx          As Integer

Private Sub btns_Click(Index As Integer)
     If Index = 0 Then
          Call GenerateReport(SelInx, MonthView.Value)
     ElseIf Index = 1 Then
          Unload Me
     End If
End Sub

Private Sub Form_Load()
     MonthView.Value = Date
     
     SelInx = 1
End Sub

Private Sub Option1_Click(Index As Integer)
     SelInx = Index + 1
End Sub
