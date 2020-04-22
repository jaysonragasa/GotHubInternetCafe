VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{E6E03A98-C7DC-4FCE-800D-724A332410A9}#1.0#0"; "LaVolpeButtons.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Got HUB? Internet and Service Center - Server Side"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   300
   ClientWidth     =   11520
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11520
   Begin VB.PictureBox Frame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2040
      Index           =   1
      Left            =   3960
      ScaleHeight     =   136
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   495
      TabIndex        =   57
      Top             =   6105
      Width           =   7425
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   23
      Top             =   8475
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15134
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            TextSave        =   "2/17/2007"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckBlocker 
      Index           =   0
      Left            =   4530
      Top             =   7455
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox TopFrame 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   768
      TabIndex        =   36
      Top             =   0
      Width           =   11520
      Begin VB.PictureBox picMenuFrame 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   0
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   789
         TabIndex        =   37
         Top             =   1335
         Width           =   11835
         Begin LaVolpeButtons.lvButtons_H btnMenu 
            Height          =   465
            Index           =   0
            Left            =   0
            TabIndex        =   47
            ToolTipText     =   "Change PC Name, IP Address"
            Top             =   0
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   820
            Caption         =   "Workstation Manager"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16777215
            cFHover         =   16777215
            cBhover         =   0
            Focus           =   0   'False
            cGradient       =   0
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmMain.frx":57E2
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frmMain.frx":5D7C
         End
         Begin LaVolpeButtons.lvButtons_H btnMenu 
            Height          =   465
            Index           =   1
            Left            =   2265
            TabIndex        =   48
            ToolTipText     =   "View all services, prices. Edit them with  any amount you want"
            Top             =   0
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   820
            Caption         =   "Services Manager"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16777215
            cFHover         =   16777215
            cBhover         =   0
            Focus           =   0   'False
            cGradient       =   0
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmMain.frx":6A56
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frmMain.frx":6FF0
         End
         Begin LaVolpeButtons.lvButtons_H btnMenu 
            Height          =   465
            Index           =   2
            Left            =   4275
            TabIndex        =   52
            ToolTipText     =   "View all services, prices. Edit them with  any amount you want"
            Top             =   0
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   820
            Caption         =   "Report"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16777215
            cFHover         =   16777215
            cBhover         =   0
            Focus           =   0   'False
            cGradient       =   0
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmMain.frx":7CCA
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frmMain.frx":8264
         End
      End
      Begin LaVolpeButtons.lvButtons_H lvBtns 
         Height          =   315
         Index           =   0
         Left            =   8355
         TabIndex        =   55
         Top             =   885
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         CapAlign        =   2
         BackStyle       =   5
         Shape           =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 2"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777215
         Focus           =   0   'False
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":8F3E
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin LaVolpeButtons.lvButtons_H lvBtns 
         Height          =   315
         Index           =   1
         Left            =   7995
         TabIndex        =   56
         Top             =   885
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         CapAlign        =   2
         BackStyle       =   5
         Shape           =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777215
         Focus           =   0   'False
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":953D
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.Image Image2 
         Height          =   1800
         Index           =   2
         Left            =   8355
         Picture         =   "frmMain.frx":99D2
         Top             =   -60
         Width           =   4065
      End
      Begin VB.Image Image2 
         Height          =   1800
         Index           =   1
         Left            =   8130
         Picture         =   "frmMain.frx":21894
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   195
      End
      Begin VB.Image Image2 
         Height          =   1800
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":22B96
         Top             =   0
         Width           =   7890
      End
   End
   Begin VB.PictureBox picOist 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   10890
      ScaleHeight     =   375
      ScaleWidth      =   525
      TabIndex        =   49
      Top             =   1650
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   4035
      Top             =   7455
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox WSM_Frame 
      BackColor       =   &H80000005&
      Height          =   4185
      Left            =   3960
      ScaleHeight     =   4125
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   1770
      Width           =   6645
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5970
         Top             =   3450
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":51078
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":51612
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":51BAC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin GoHUB_Server.InfoHeader ihFloatCngWS 
         Height          =   2190
         Left            =   2520
         TabIndex        =   32
         Top             =   1185
         Visible         =   0   'False
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   3863
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   16777215
         FrameBorderColor=   -2147483632
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
         Caption         =   "Change Workstation"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMain.frx":51F46
         EdgeRadiusSize  =   32
         RoundEdgeStyle  =   1
         Begin VB.TextBox txSelWS 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   705
            Width           =   1710
         End
         Begin VB.CommandButton CW_Btns 
            Caption         =   "&Set"
            Height          =   300
            Index           =   1
            Left            =   3690
            TabIndex        =   46
            Top             =   1800
            Width           =   570
         End
         Begin VB.CommandButton CW_Btns 
            Caption         =   "&Cancel"
            Height          =   300
            Index           =   0
            Left            =   2805
            TabIndex        =   45
            Top             =   1800
            Width           =   810
         End
         Begin VB.ListBox lsMovWS 
            Height          =   885
            IntegralHeight  =   0   'False
            Left            =   2145
            TabIndex        =   44
            Top             =   705
            Width           =   2115
         End
         Begin VB.Image Image1 
            Height          =   15
            Index           =   2
            Left            =   60
            Picture         =   "frmMain.frx":524E0
            Stretch         =   -1  'True
            Top             =   1710
            Width           =   4140
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Move To This Workstation"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   43
            Top             =   480
            Width           =   1860
         End
         Begin VB.Label lblCW 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$name$"
            Height          =   195
            Index           =   0
            Left            =   555
            TabIndex        =   35
            Top             =   1350
            Width           =   570
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   1
            Left            =   210
            Picture         =   "frmMain.frx":5297A
            Top             =   1335
            Width           =   240
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Loged in User:"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   34
            Top             =   1125
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Workstation"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   33
            Top             =   480
            Width           =   1530
         End
      End
      Begin GoHUB_Server.InfoHeader ihFloatTLFrame 
         Height          =   1395
         Left            =   345
         TabIndex        =   16
         Top             =   1170
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   2461
         InfoHeaderStyle =   1
         HeaderHeight    =   21
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
         Caption         =   "Set Time Limit"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMain.frx":52D04
         EdgeRadiusSize  =   32
         RoundEdgeStyle  =   1
         Begin VB.CommandButton TL_Btns 
            Caption         =   "&Cancel"
            Height          =   300
            Index           =   1
            Left            =   735
            TabIndex        =   22
            Top             =   990
            Width           =   810
         End
         Begin VB.CommandButton TL_Btns 
            Caption         =   "&Set"
            Height          =   300
            Index           =   0
            Left            =   1605
            TabIndex        =   21
            Top             =   990
            Width           =   570
         End
         Begin VB.TextBox txTL 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1650
            TabIndex        =   20
            Text            =   "00"
            Top             =   450
            Width           =   420
         End
         Begin VB.TextBox txTL 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   555
            TabIndex        =   19
            Text            =   "00"
            Top             =   450
            Width           =   420
         End
         Begin VB.Image Image1 
            Height          =   15
            Index           =   0
            Left            =   75
            Picture         =   "frmMain.frx":5309E
            Stretch         =   -1  'True
            Top             =   900
            Width           =   2115
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minute"
            Height          =   195
            Index           =   1
            Left            =   1125
            TabIndex        =   18
            Top             =   495
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hour"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   480
            Width           =   345
         End
      End
      Begin VB.PictureBox picFloatShadow 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   1
         Left            =   4530
         ScaleHeight     =   840
         ScaleWidth      =   795
         TabIndex        =   24
         Top             =   1485
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5610
         Top             =   225
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":53538
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":538D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":53C6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":54006
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":545A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":54B3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":54ED4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   540
         Left            =   120
         TabIndex        =   15
         Top             =   585
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   953
         ButtonWidth     =   1402
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Log In"
               Object.ToolTipText     =   "Login selected workstation"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Log Out"
               Object.ToolTipText     =   "Logout selected workstation"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Time Limit"
               Object.ToolTipText     =   "Set time limit on selected workstation"
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Change"
               Object.ToolTipText     =   "Change loged in workstaion to another workstation"
               ImageIndex      =   4
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cancel"
               Object.ToolTipText     =   "cancel user usage"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin GoHUB_Server.InfoHeader IH 
         Height          =   570
         Index           =   0
         Left            =   345
         TabIndex        =   14
         Top             =   90
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   1005
         InfoHeaderStyle =   0
         HeaderHeight    =   21
         FrameBackColor  =   16777215
         FrameBorderColor=   10070188
         GradientStyle   =   1
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
         Caption         =   $"frmMain.frx":5546E
         MultiLine       =   -1  'True
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMain.frx":554C0
         EdgeRadiusSize  =   6
         RoundEdgeStyle  =   6
      End
      Begin MSComctlLib.ListView lvClients 
         Height          =   2100
         Left            =   870
         TabIndex        =   13
         ToolTipText     =   "Select any PC Letter and click any buttons on the toolbar"
         Top             =   1635
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3704
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox Frame 
      Align           =   3  'Align Left
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   6675
      Index           =   0
      Left            =   0
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   1800
      Width           =   3675
      Begin GoHUB_Server.InfoHeader IH 
         Height          =   2985
         Index           =   1
         Left            =   75
         TabIndex        =   2
         Top             =   75
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   5265
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   16777215
         FrameBorderColor=   -2147483646
         GradientStyle   =   0
         LeftColor       =   6956042
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
         ForeColor       =   -2147483639
         Caption         =   "Services"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMain.frx":55A5A
         EdgeRadiusSize  =   8
         RoundEdgeStyle  =   1
         Begin VB.PictureBox Float_picQuan 
            Height          =   690
            Left            =   1050
            ScaleHeight     =   630
            ScaleWidth      =   1350
            TabIndex        =   25
            Top             =   900
            Visible         =   0   'False
            Width           =   1410
            Begin VB.TextBox txQuan 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   135
               Locked          =   -1  'True
               TabIndex        =   26
               Text            =   "0"
               Top             =   285
               Width           =   840
            End
            Begin ComCtl2.UpDown UpDown1 
               Height          =   285
               Left            =   945
               TabIndex        =   27
               Top             =   285
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   327681
               BuddyControl    =   "txQuan"
               BuddyDispid     =   196627
               OrigLeft        =   780
               OrigTop         =   285
               OrigRight       =   1035
               OrigBottom      =   450
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "Wingdings 2"
                  Size            =   12
                  Charset         =   2
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   255
               Left            =   1110
               MouseIcon       =   "frmMain.frx":55FF4
               MousePointer    =   99  'Custom
               TabIndex        =   29
               ToolTipText     =   "update user services"
               Top             =   0
               Width           =   225
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000010&
               Caption         =   " Quantity"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   225
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   1350
            End
         End
         Begin MSComctlLib.ListView lvServList 
            Height          =   1770
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "select any services for the selected client by clicking on the  checkbox found right beside the service name."
            Top             =   765
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   3122
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ColHdrIcons     =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Services"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Per Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Quantity"
               Object.Width           =   2540
            EndProperty
         End
         Begin GoHUB_Server.AdvanceLabelControl Srv_ALbl 
            Height          =   270
            Index           =   0
            Left            =   150
            TabIndex        =   9
            ToolTipText     =   "Walkthrough Service is for blah-blah-blah.."
            Top             =   2640
            Width           =   1785
            _ExtentX        =   3149
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
            Caption         =   "Walkthrough Service"
            NormalIcon      =   "frmMain.frx":56CBE
            HasIcon         =   -1  'True
            AutoResize      =   -1  'True
            OnHoverStyle    =   3
            OnHoverForeColor=   -2147483635
         End
         Begin VB.Image Image1 
            Height          =   30
            Index           =   1
            Left            =   30
            Picture         =   "frmMain.frx":57258
            Stretch         =   -1  'True
            Top             =   2580
            Width           =   3315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select client services"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   3
            Top             =   495
            Width           =   1485
         End
      End
      Begin GoHUB_Server.InfoHeader IH 
         Height          =   1740
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   2190
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   3069
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   -2147483643
         FrameBorderColor=   -2147483646
         GradientStyle   =   0
         LeftColor       =   6956042
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
         ForeColor       =   -2147483639
         Caption         =   "Client Tools"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMain.frx":576F2
         EdgeRadiusSize  =   12
         RoundEdgeStyle  =   6
         Begin GoHUB_Server.AdvanceLabelControl CT_ALbl 
            Height          =   270
            Index           =   1
            Left            =   150
            TabIndex        =   11
            Tag             =   "Monitors the workstations screen, and more!"
            Top             =   705
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
            BackColor       =   -2147483643
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Monitor Client"
            NormalIcon      =   "frmMain.frx":57C8C
            HasIcon         =   -1  'True
            AutoResize      =   -1  'True
            OnHoverStyle    =   3
            OnHoverForeColor=   -2147483635
         End
         Begin GoHUB_Server.AdvanceLabelControl CT_ALbl 
            Height          =   270
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Tag             =   "send message to the user."
            Top             =   420
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            BackColor       =   -2147483643
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Send Message"
            NormalIcon      =   "frmMain.frx":58026
            HasIcon         =   -1  'True
            AutoResize      =   -1  'True
            OnHoverStyle    =   3
            OnHoverForeColor=   -2147483635
         End
         Begin VB.Image Image1 
            Height          =   30
            Index           =   3
            Left            =   120
            Picture         =   "frmMain.frx":583C0
            Stretch         =   -1  'True
            Top             =   1065
            Width           =   3285
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   75
            Picture         =   "frmMain.frx":5885A
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label CT_lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Client Tools, "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   390
            TabIndex        =   5
            Top             =   1185
            Width           =   2400
         End
      End
      Begin GoHUB_Server.InfoHeader IH 
         Height          =   1770
         Index           =   4
         Left            =   60
         TabIndex        =   6
         Top             =   4875
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   3122
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   -2147483643
         FrameBorderColor=   -2147483646
         GradientStyle   =   0
         LeftColor       =   6956042
         RightColor      =   0
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
         ForeColor       =   -2147483639
         Caption         =   "Overall Status"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMain.frx":58BE4
         EdgeRadiusSize  =   6
         RoundEdgeStyle  =   6
         Begin VB.Label lblOA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ": 0"
            Height          =   195
            Index           =   2
            Left            =   2880
            TabIndex        =   54
            Top             =   840
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PC Connected"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   53
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblOA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ": 0"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   31
            Top             =   645
            Width           =   195
         End
         Begin VB.Label lblOA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ": 0"
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   30
            Top             =   450
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PC Not Used"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   8
            Top             =   645
            Width           =   900
         End
         Begin VB.Label ads 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PC Used"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   7
            Top             =   450
            Width           =   600
         End
      End
      Begin GoHUB_Server.InfoHeader IH 
         Height          =   1500
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   3585
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   2646
         InfoHeaderStyle =   1
         HeaderHeight    =   21
         FrameBackColor  =   -2147483643
         FrameBorderColor=   -2147483646
         GradientStyle   =   0
         LeftColor       =   6956042
         RightColor      =   0
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
         ForeColor       =   -2147483639
         Caption         =   "Information"
         MultiLine       =   0   'False
         Alignment       =   0
         HasIcon         =   -1  'True
         Picture         =   "frmMain.frx":5917E
         EdgeRadiusSize  =   6
         RoundEdgeStyle  =   6
         Begin VB.Timer Timer1 
            Interval        =   6000
            Left            =   60
            Top             =   960
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Height          =   1050
            Left            =   645
            TabIndex        =   50
            Top             =   375
            Width           =   2805
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   90
            Picture         =   "frmMain.frx":59518
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.PictureBox picShadowIH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1065
         Index           =   0
         Left            =   495
         ScaleHeight     =   1065
         ScaleWidth      =   3195
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.PictureBox picShadowIH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1065
         Index           =   1
         Left            =   465
         ScaleHeight     =   1065
         ScaleWidth      =   3195
         TabIndex        =   40
         Top             =   2580
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.PictureBox picShadowIH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1065
         Index           =   3
         Left            =   840
         ScaleHeight     =   1065
         ScaleWidth      =   3195
         TabIndex        =   42
         Top             =   5460
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.PictureBox picShadowIH 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1065
         Index           =   2
         Left            =   405
         ScaleHeight     =   1065
         ScaleWidth      =   3195
         TabIndex        =   41
         Top             =   3720
         Visible         =   0   'False
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type
'
'Private Type MINMAXINFO
'    ptReserved As POINTAPI
'    ptMaxSize As POINTAPI
'    ptMaxPosition As POINTAPI
'    ptMinTrackSize As POINTAPI
'    ptMaxTrackSize As POINTAPI
'End Type
'
'Private Const WM_GETMINMAXINFO = &H24
'Private Const WM_ENTERSIZEMOVE = &H231
'Private Const WM_EXITSIZEMOVE = &H232

Dim FiveCnt                        As Integer
Dim Desktop                        As DesktopArea
Dim TmrTemp                        As Integer
Dim m_Sizing                       As Boolean

Dim SelServIndex                   As Integer

Dim ttlsnt                         As Long

'Implements ISubclass
'Private m_emr As EMsgResponse
Private WithEvents Timer           As ccrpTimer
Attribute Timer.VB_VarHelpID = -1

'Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
'    m_emr = RHS
'End Property
'
'Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
'    Debug.Print CurrentMessage
'    m_emr = emrConsume
'    ISubclass_MsgResponse = m_emr
'End Property
'
'Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'     Dim mmiT As MINMAXINFO
'
'     Select Case iMsg
'          Case WM_ENTERSIZEMOVE
'               m_Sizing = True
'          Case WM_EXITSIZEMOVE
'               m_Sizing = False
'
'               Call InitGUI
'               'Call picMenuFrame_Resize
'               'Call Frame_Resize(0)
'               'Call Frame_Resize(1)
'          Case WM_GETMINMAXINFO
'              ' Copy parameter to local variable for processing
'              CopyMemory mmiT, ByVal lParam, LenB(mmiT)
'
'              ' Minimium width and height for sizing
'              mmiT.ptMinTrackSize.x = 800
'              mmiT.ptMinTrackSize.y = 600
'
'              ' Copy modified results back to parameter
'              CopyMemory ByVal lParam, mmiT, LenB(mmiT)
'     End Select
'End Function

Private Sub btnMenu_Click(Index As Integer)
     If Index = 0 Then
          frmWSMngr.Show vbModal, Me
     ElseIf Index = 1 Then
          frmServMngr.Show vbModal, Me
     ElseIf Index = 2 Then
          frmReport.Show vbModal, Me
     End If
End Sub

Private Sub CT_ALbl_Click(Index As Integer)
     If frmMain.sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State <> sckConnected Then
          CT_ALbl(0).Enabled = False: CT_ALbl(1).Enabled = False
          
          MsgBox "workstation not connected", vbExclamation
          
          Exit Sub
     End If
     
     If Index = 0 Then        ' chat
          frmChat.Show vbModal, Me
     ElseIf Index = 1 Then    ' montior
          
          frmMonitor.Show vbModal, Me
     End If
End Sub

Private Sub CT_ALbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     CT_lblinfo.Caption = CT_ALbl(Index).Tag
End Sub

Private Sub CW_Btns_Click(Index As Integer)
     Dim nInx       As Integer
     Dim i          As Integer
     Dim c          As Integer
     Dim Item       As ListItem
     Dim temp()     As String
     Dim oldIndx    As Integer
     
     If Index = 0 Then
          Call HideAllFloatingControl(Me)
          Toolbar.Buttons(5).Value = tbrUnpressed
          
          Exit Sub
     End If
     
     oldIndx = ItmDetails.ItemIndex
     
     For i = 1 To lvClients.ListItems.Count
          If lvClients.ListItems(i).Text = lsMovWS.List(lsMovWS.ListIndex) Then
               For c = 1 To lvClients.ColumnHeaders.Count - 1
                    lvClients.ListItems(i).SubItems(c) = lvClients.ListItems(ItmDetails.ItemIndex).SubItems(c)
                    lvClients.ListItems(ItmDetails.ItemIndex).SubItems(c) = vbNullString
               Next c
               
               lvClients.ListItems(ItmDetails.ItemIndex).Tag = 0
               Set Item = lvClients.ListItems(i)
               Item.Tag = 1
               Item.Selected = True
               
               temp = Split(lvClients.ListItems(i).Key, "|")
               Call ChangeWorkstation(ItmDetails.IPNumber, temp(1))
               
               temp = Split(lvClients.ListItems(i).Key, "|")
               With ItmDetails
                    .ItemIndex = i
                    .ItemName = lvClients.ListItems(i).Text
                    .ClientPCID = temp(0)
                    .IPNumber = temp(1)
                    .IsConnected = True
                    .IsLogedIn = True
               End With
               
               If i <> 0 Then
                    If sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected Then
                         sckServer(WSAddVars(i).WinsockIndex).SendData CONN_REQUESTSTATUS & "|" & _
                                                                       lvClients.ListItems(i).SubItems(1) & "|" & _
                                                                       lvClients.ListItems(i).SubItems(2) & "|" & _
                                                                       Left$(lvClients.ListItems(i).SubItems(4), 8)
                         DoEvents
                                                                  
                         sckServer(WSAddVars(oldIndx).WinsockIndex).SendData CONN_CANCEL
                         DoEvents
                    
                         Sleep 500
                    End If
               End If
     
               Exit For
          End If
     Next i
     
     Call HideAllFloatingControl(Me)
     Toolbar.Buttons(5).Value = tbrUnpressed
End Sub

Private Sub ddWS_Change()

End Sub

Private Sub Form_Initialize()
     Dim desk As New DesktopArea
     
     desk.PositionForm Me, H_FULL, V_FULL
     'Move 0, 0, tX(1024), tY(768)
     
'     AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO
'     AttachMessage Me, Me.hwnd, WM_ENTERSIZEMOVE
'     AttachMessage Me, Me.hwnd, WM_EXITSIZEMOVE

     ' [Setup CCRP High Performance Timer Control]
     ' -------------------------------------------------------------------------------------------------------------
     ' -------------------------------------------------------------------------------------------------------------
     Set Timer = New ccrpTimer
     Timer.interval = 1000
     Timer.Enabled = False
     ' -------------------------------------------------------------------------------------------------------------
     ' -------------------------------------------------------------------------------------------------------------
     
     ' enable XP control support
     Call FixThemeSupport(Controls)
     
     ' make sure all control are in position
     'Call Form_Resize
     
     ' setup lvClients ColumnHeaders
     Call InitColumns
     
     Call GetWorkStationNames(lvClients)
     Call GetServices(lvServList)
     
     Call InitGUI
     
     Call Listen
End Sub

Private Sub InitGUI()
     Hide
     
     BackColor = RGB(255, 255, 255) 'RGB(199, 211, 247)
     WSM_Frame.BackColor = BackColor
     
     ' create gradient effect on the left panel
     'Call Gradient(Frame(0), RGB(10, 36, 106), RGB(5, 19, 56), "VERTICAL")
     Call Gradient(Frame(0), RGB(122, 161, 230), vbWhite, "VERTICAL")
     Call Gradient(picMenuFrame, RGB(106, 143, 217), RGB(51, 102, 204), "VERTICAL")
     'Call Gradient(Frame(1), RGB(106, 143, 217), RGB(51, 102, 204), "VERTICAL")
     Call Gradient(Frame(1), RGB(255, 255, 255), SystemColorConstants.vb3DFace, "VERTICAL")
     'Call Gradient(picMenuFrame, RGB(199, 211, 247), RGB(51, 102, 204), "VERTICAL")
     'Call Gradient(TopFrame, RGB(255, 255, 255), RGB(199, 211, 247), "VERTICAL")
     
     ' [Set Control Styles]
     ' -------------------------------------------------------------------------------------------------------------
     ' -------------------------------------------------------------------------------------------------------------
     Call InitInfoHeaders
     Call SetupMenuButtons
     
     WSM_Frame.BorderStyle = 0
     Frame(0).BorderStyle = 0
     
     Call Flatten_ListView_ColumnButton(lvClients)
     Call Flatten_ListView_ColumnButton(lvServList)
     
     Call StyleStaticEdge(lvClients)
     Call StyleStaticEdge(lvServList)
     Call DrawModalFrame(Float_picQuan.hwnd)
     
     ' -------------------------------------------------------------------------------------------------------------
     ' -------------------------------------------------------------------------------------------------------------
     
     Image2(0).Move 0, 0
     Image2(2).Move TopFrame.ScaleWidth - Image2(2).Width, 0
     Image2(1).Move Image2(0).Left + Image2(0).Width, 0
     Image2(1).Width = Image2(2).Left - Image2(1).Left
     
     Unload frmSplash
     
     Me.Show: DoEvents
     
     If Not IsIDE Then
          Call DrawShadow(Me, TopFrame, picOist, 2)
          Call DrawIHPanelsShadow(1)
     End If
     
     Timer.Enabled = True
End Sub

Sub SetupMenuButtons()
     Dim i          As Integer
     
     For i = 0 To btnMenu.Count - 1
          btnMenu(i).BackColor = RGB(106, 143, 217)
          btnMenu(i).GradientColor = RGB(51, 102, 204)
          btnMenu(i).ForeColor = vbWhite
     Next i
End Sub

Sub InitColumns()
     Dim i          As Integer
     
     With lvClients
          .ColumnHeaders.Add , , "Workstation"
          .ColumnHeaders.Add , , "User Name"
          .ColumnHeaders.Add , , "LogIn Time"
          .ColumnHeaders.Add , , "LogOut Time", , vbCenter
          .ColumnHeaders.Add , , "Time Used"
          .ColumnHeaders.Add , , vbNullString '"Comment"
          .ColumnHeaders.Add , , vbNullString '"Currently Accessing"
          
          .View = lvwReport
          .GridLines = True
          .FullRowSelect = True
          .LabelEdit = lvwManual
          
          Set .SmallIcons = ImageList2
     End With
     
     Call AutoResizeListView(lvClients)
End Sub

Sub InitInfoHeaders()
     Dim i          As Integer
     
     IH(0).EdgeRadiusSize = 12
     IH(0).LeftColor = vbWhite
     IH(0).RightColor = SystemColorConstants.vb3DFace
     IH(0).ForeColor = SystemColorConstants.vbButtonText
     
     For i = 1 To 4
          IH(i).EdgeRadiusSize = 6
          IH(i).LeftColor = vbWhite
          IH(i).RightColor = RGB(199, 211, 247)
          IH(i).ForeColor = RGB(33, 93, 198)
          IH(i).FrameBorderColor = RGB(199, 211, 247)
          IH(i).FrameBackColor = vbWhite
          IH(i).RoundEdgePosition = REP_Top
          DoEvents
     Next i
     
     IH(4).LeftColor = RGB(51, 102, 204)
     IH(4).RightColor = RGB(98, 116, 213)
     IH(4).ForeColor = vbWhite
     IH(4).FrameBorderColor = RGB(98, 116, 213)
     
     IH(1).Move 5, 8
     IH(1).Width = Frame(0).ScaleWidth - IH(1).Left - 5
     
     For i = 2 To IH.Count - 1
          IH(i).Move IH(i - 1).Left, IH(i - 1).Top + IH(i - 1).Height + 5
          IH(i).Width = IH(i - 1).Width
     Next i
     
     IH(4).Top = Frame(0).ScaleHeight - IH(4).Height - 5
     
     ihFloatTLFrame.LeftColor = SystemColorConstants.vb3DFace
     ihFloatTLFrame.RightColor = SystemColorConstants.vbWindowBackground
     ihFloatTLFrame.ForeColor = SystemColorConstants.vbButtonText
     ihFloatTLFrame.FrameBorderColor = SystemColorConstants.vb3DShadow
     
     ihFloatCngWS.LeftColor = SystemColorConstants.vb3DFace
     ihFloatCngWS.RightColor = SystemColorConstants.vbWindowBackground
     ihFloatCngWS.ForeColor = SystemColorConstants.vbButtonText
     ihFloatCngWS.FrameBorderColor = SystemColorConstants.vb3DShadow
     
     Frame(0).ScaleMode = vbTwips
     
     DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Dim i          As Integer
     
     Timer.Enabled = False
     Set Timer = Nothing
     DoEvents
     
'     DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
'     DetachMessage Me, Me.hwnd, WM_ENTERSIZEMOVE
'     DetachMessage Me, Me.hwnd, WM_EXITSIZEMOVE
     
     For i = 0 To sckBlocker.Count - 1
          sckBlocker(i).Close
     Next i
     
     For i = 0 To sckServer.Count - 1
          sckServer(i).Close
     Next i
End Sub

Private Sub DrawIHPanelsShadow(ByVal Opacity As Integer)
     Call DrawShadow_IC(Me, Frame(0), IH(1), picShadowIH(0), Opacity)
     Call DrawShadow_IC(Me, Frame(0), IH(2), picShadowIH(1), Opacity)
     Call DrawShadow_IC(Me, Frame(0), IH(3), picShadowIH(2), Opacity)
     Call DrawShadow_IC(Me, Frame(0), IH(4), picShadowIH(3), Opacity)
End Sub

Private Sub Label4_Click()         ' the X label from Quantity Frame
     ' update Quantity Column
     lvServList.ListItems(SelServListIndx).SubItems(3) = CInt(txQuan.Text)
     
     If CInt(txQuan.Text) <> 0 Then
          lvServList.ListItems(SelServListIndx).Checked = True
          
          Call AddClientServices(ItmDetails.IPNumber, lvServList)
     Else
          lvServList.ListItems(SelServListIndx).Checked = False
     End If
     
     Call HideAllFloatingControl(Me)
End Sub

Private Sub lvbtns_Click(Index As Integer)
     If Index = 0 Then   ' Exit
          Unload Me
     ElseIf Index = 1 Then
          Me.WindowState = vbMinimized
     End If
End Sub

Private Sub lvClients_ItemClick(ByVal Item As MSComctlLib.ListItem)
     On Error Resume Next
     
     Dim temp()          As String
     Dim i               As Integer
     
     Dim cTime           As String
     Dim stat            As String
     Dim INet            As Currency
     Dim PMin            As Currency
     Dim TtlInetAmnt     As Currency
     Dim TtlServ         As Currency
     
     temp = Split(Item.Key, "|")
     
     Call GetClientServices(lvServList, temp(1))
     
     Call HideAllFloatingControl(Me)
     Toolbar.Buttons(3).Value = tbrUnpressed
     Toolbar.Buttons(5).Value = tbrUnpressed
     
     With ItmDetails
          .ItemIndex = Item.Index
          .ItemName = Item.Text
          .ClientPCID = temp(0)
          .IPNumber = temp(1)
          .IsConnected = IIf(sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected, True, False)
          .IsLogedIn = CBool(Item.Tag)
          
          lvServList.Enabled = .IsLogedIn
          
          CT_ALbl(0).Enabled = .IsConnected 'IIf(Item.SmallIcon = 1, True, False)
          CT_ALbl(1).Enabled = .IsConnected 'IIf(Item.SmallIcon = 1, True, False)
          
          StatusBar.Panels(1).Text = "Selected Workstation: " + .ItemName + " [" + .IPNumber + "]"
          
          If Not .IsConnected Then
               'MsgBox "Workstation not connected.", vbExclamation
               
               Toolbar.Buttons(1).Enabled = True
               For i = 2 To Toolbar.Buttons.Count
                    Toolbar.Buttons(i).Enabled = False
               Next i
               
               If lvClients.ListItems(Item.Index).SubItems(1) <> vbNullString Then
                    Toolbar.Buttons(1).Enabled = False
                    
                    For i = 2 To Toolbar.Buttons.Count
                         Toolbar.Buttons(i).Enabled = True
                    Next i
               End If
          ElseIf Not .IsLogedIn Then
               Toolbar.Buttons(1).Enabled = True
                    
               For i = 2 To Toolbar.Buttons.Count
                    Toolbar.Buttons(i).Enabled = False
               Next i
          End If
          If lvClients.ListItems(Item.Index).SubItems(1) <> vbNullString Then
               Toolbar.Buttons(1).Enabled = False
               For i = 2 To Toolbar.Buttons.Count
                    Toolbar.Buttons(i).Enabled = True
               Next i
               
               ' users inet and serv amount status
               If lvServList.ListItems(1).Checked = True Then
                    INet = lvServList.ListItems(1).SubItems(1)
               ElseIf lvServList.ListItems(2).Checked = True Then
                    INet = lvServList.ListItems(2).SubItems(1)
               End If
               
               For i = 3 To lvServList.ListItems.Count
                    If lvServList.ListItems(i).Checked = True Then
                         TtlServ = TtlServ + lvServList.ListItems(i).SubItems(1) * lvServList.ListItems(i).SubItems(3)
                    End If
               Next i
               
               PMin = INet / 60
               
               cTime = lvClients.ListItems(Item.Index).SubItems(2)
               cTime = DateDiff("s", cTime, Time)
               cTime = cTime \ 60
               
               TtlInetAmnt = cTime * PMin
               
               stat = "User: " + lvClients.ListItems(Item.Index).SubItems(1) + vbCrLf
               stat = stat + "LogIn Date: " & Date & vbCrLf
               stat = stat + "Time Used: " + lvClients.ListItems(Item.Index).SubItems(4) + vbCrLf
               stat = stat + "Iternet Usage Amount: " & FormatNumber(TtlInetAmnt, 2) & vbCrLf
               stat = stat + "Total Services: " & TtlServ
               
               lblInfo.Caption = stat
          End If
     End With
End Sub

Private Sub lvServList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     ' we dont want to have a both check marks on Internet Service and Internet with WebCam Service.
     ' so heres how to prevent that
     ' --------------------------------------------------------------------------------------------------------------
     If Item.Index = 1 Or Item.Index = 2 Then
          lvServList.ListItems(1).Checked = False
          lvServList.ListItems(2).Checked = False
     
          Item.Checked = True
          
          Call ChangeInternetService(ItmDetails.IPNumber, Item.Key)
     End If
     ' --------------------------------------------------------------------------------------------------------------

     ' dont inlcude "Internet Service"
     If Item.Index = 1 Or Item.Index = 2 Then Exit Sub
     
     If Item.Checked = True Then
          Float_picQuan.Visible = False
          Float_picQuan.Top = lvServList.Top
          Float_picQuan.Left = lvServList.Width - Float_picQuan.Width
          Float_picQuan.Visible = True
          
          txQuan.Text = Item.SubItems(3)
          
          DoEvents
          
          SelectedService = Item.Key: SelServListIndx = Item.Index
     Else
          Call RemoveSelectedService(ItmDetails.IPNumber, Item.Key)
          
          SelectedService = vbNullString
          SelServListIndx = 0
          Call HideAllFloatingControl(Me)
     End If
     
     SelServListIndx = Item.Index
End Sub

Private Sub lvServList_ItemClick(ByVal Item As MSComctlLib.ListItem)
     If Item.Index = 1 Or Item.Index = 2 Then Exit Sub
     
     SelectedService = Item.Key: SelServListIndx = Item.Index
     
     ' if the selected item already have an quantity value then show the picQuan for Quantity Update
     If Item.SubItems(3) <> 0 Then
          Float_picQuan.Visible = False
          Float_picQuan.Top = lvServList.Top
          Float_picQuan.Left = lvServList.Width - Float_picQuan.Width
          Float_picQuan.Visible = True
          txQuan.Text = Item.SubItems(3)
     End If
End Sub

Private Sub Srv_ALbl_Click(Index As Integer)
     If Index = 0 Then
          frmWalkthroughServ.Show vbModal, Me
     End If
End Sub

Private Sub Timer_Timer(ByVal Milliseconds As Long)
     On Error Resume Next
     
     Dim i          As Integer
     Dim temp       As String
     Dim TimeUsed   As Variant
     
     Dim used       As Integer
     Dim cnctd      As Integer
     Dim tmrUsed    As String
     
     Dim sParse()        As String
     
     If FiveCnt = 5 Then FiveCnt = 0
     
     FiveCnt = FiveCnt + 1
     
     lvClients.Refresh
     
     If lvClients.ListItems.Count > 0 Then
          For i = 1 To lvClients.ListItems.Count
               If lvClients.ListItems(i).Tag = 1 Then  ' Tag=1, workstation is in use, Tag=0 wala gumagamit
                    temp = Format(TimeValue(Time) - TimeValue(lvClients.ListItems(i).SubItems(2)), "hh:mm:ss")
                    
                    TimeUsed = Split(temp, ":")
                    
                    tmrUsed = TimeUsed(0) & "h, " & TimeUsed(1) & "m, " & TimeUsed(2) & "s" 'Left$(TimeUsed(2), Len(TimeUsed(2)) - 3) & "s"
                    
                    ' for every change of minute
                    If TimeUsed(1) <> TmrTemp Then
                         If sckServer(WSAddVars(i).WinsockIndex).State = sckConnected Then
                              sckServer(WSAddVars(i).WinsockIndex).SendData "TIMER|" & lvClients.ListItems(i).SubItems(2) & "|" + Left$(tmrUsed, 8)
                              
                              DoEvents
                         End If
                    End If
                    
                    If lvClients.ListItems(i).SubItems(3) <> vbNullString Then
                         If FormatDateTime(lvClients.ListItems(i).SubItems(3), vbLongTime) = FormatDateTime(Time, vbLongTime) Then
                              temp = lvClients.ListItems(i).SubItems(1)
                              lvClients.ListItems(i).Tag = 0
                              lvClients.ListItems(i).SmallIcon = 3
                              lvClients.ListItems(i).SubItems(1) = temp & ": TIMES UP"
                         
                              'sckServer(WSAddVars(i).WinsockIndex).SendData CONN_TIMESUP
                              
                              DoEvents
                         End If
                    End If
                    
                    lvClients.ListItems(i).SubItems(4) = tmrUsed
                    TmrTemp = TimeUsed(1)
                    
                    used = used + 1
                    
                    If FiveCnt = 5 Then Call SetTemporaryRecord(i, False)
               End If
               
               If sckServer(WSAddVars(i).WinsockIndex).State = sckConnected Then
                    cnctd = cnctd + 1
     
               Else
                    sParse = Split(lvClients.ListItems(i).Key, "|")
                    
                    Call SetConnectComputerEffect(lvClients, sParse(1), False)
               End If
               
               If lvClients.ListItems(i).SubItems(1) = vbNullString Then
                    sParse = Split(lvClients.ListItems(i).Key, "|")
                    
                    Call GetClientStatus(lvClients, i, sParse(1))
               End If
          Next i
          
          Call AutoResizeListView(lvClients)
     End If
     
     StatusBar.Panels(2).Text = FormatDateTime(Time, vbLongTime)
     
     With OverallStat
          .OA_PCUSED = used
          .OA_PCNOTUSED = lvClients.ListItems.Count - used
          .OA_PCCONNECTED = cnctd
     
          lblOA(0).Caption = ": " & .OA_PCUSED
          lblOA(1).Caption = ": " & .OA_PCNOTUSED
          lblOA(2).Caption = ": " & .OA_PCCONNECTED
     End With
End Sub

Private Sub TL_Btns_Click(Index As Integer)
     Dim temp       As String
     Dim LIT        As String      ' Log In Time
     
     If Index = 1 Then
          Call HideAllFloatingControl(Me)
          Toolbar.Buttons(3).Value = tbrUnpressed
          
          Exit Sub
     End If
     
     LIT = lvClients.ListItems(ItmDetails.ItemIndex).SubItems(2)
     temp = txTL(0).Text & ":" & txTL(1).Text
     lvClients.ListItems(ItmDetails.ItemIndex).SubItems(3) = FormatDateTime(TimeValue(LIT) + TimeValue(temp), vbLongTime)
     
     Call AutoResizeListView(lvClients)
     
     ItmDetails.ItemIndex = 0
     Call HideAllFloatingControl(Me)
     Toolbar.Buttons(3).Value = tbrUnpressed
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
     Dim i          As Integer
     Dim c          As Integer
     Dim ask        As Long
     
     Call HideAllFloatingControl(Me)
     Toolbar.Buttons(3).Value = tbrUnpressed
     Toolbar.Buttons(5).Value = tbrUnpressed
     
     If ItmDetails.ItemIndex = 0 Then
          
          Exit Sub
     End If
     
     Timer.Enabled = False
     
     If ItmDetails.ItemIndex <> 0 Then
          If Button.Index = 1 Then      ' login
               ask = MsgBox("Are you sure you want to Login this workstation", vbYesNo + vbInformation, "Confirm")
               
               If ask = vbYes Then Call LogIn
          ElseIf Button.Index = 2 Then  ' logout
               ask = MsgBox("Are you sure you want to Logout the current user", vbYesNo + vbInformation, "Confirm")
               
               If ask = vbYes Then Call LogOut
          ElseIf Button.Index = 3 Then  ' set time limit
               Call TimeLimit(Button)
          ElseIf Button.Index = 5 Then  ' change workstation
               Call Change_Workstation(Button)
          ElseIf Button.Index = 6 Then  ' cancel
               ask = MsgBox("Are you sure you want to Cancel the usage of the current user", vbYesNo + vbInformation, "Confirm")
               
               If ask = vbYes Then
                    Call CancelUsage
               End If
          ElseIf Button.Index = 8 Then
                    
          End If
     End If
     
     Timer.Enabled = True
End Sub

Private Sub TopFrame_Resize()
     picMenuFrame.Move 0, TopFrame.ScaleHeight - picMenuFrame.Height, TopFrame.ScaleWidth
     lvbtns(0).Move TopFrame.ScaleWidth - lvbtns(0).Width - 2, 2
     lvbtns(1).Move lvbtns(0).Left - lvbtns(1).Width - 2, lvbtns(0).Top
End Sub

Private Sub txTL_GotFocus(Index As Integer)
     Call AutoHighlight(txTL(Index))
End Sub

Private Sub txTL_KeyPress(Index As Integer, KeyAscii As Integer)
     If Chr$(KeyAscii) Like "[0-9]" = False Then KeyAscii = 0
End Sub

Private Sub WSM_Frame_Resize()
     IH(0).Move 0, 0, WSM_Frame.ScaleWidth
     Toolbar.Move 0, IH(0).Height, WSM_Frame.ScaleWidth
     lvClients.Move 0, Toolbar.Top + Toolbar.Height, WSM_Frame.ScaleWidth
     lvClients.Height = WSM_Frame.ScaleHeight - lvClients.Top
End Sub

Sub SetTemporaryRecord(ByVal SelIndex As Integer, ByVal Add As Boolean, Optional ByVal Services As String)
     Dim temp()          As String
     
     temp = Split(lvClients.ListItems(SelIndex).Key, "|")
     
     With TempRec
          .ClientPCID = temp(0)
          .IPNumber = temp(1)
          .PCUser = lvClients.ListItems(SelIndex).SubItems(1)
          .LogDate = Date
          .LogInTime = lvClients.ListItems(SelIndex).SubItems(2)
          .LogOutTime = lvClients.ListItems(SelIndex).SubItems(3)
          .TimeUsed = lvClients.ListItems(SelIndex).SubItems(4)
          .Services = Services
     End With
     
     Call DoTemporaryRecord(Add)
End Sub

Sub LogIn()
     Dim i          As Integer
     
     If Not ItmDetails.IsLogedIn Then
          Dim Services        As String
          Dim ask             As Long
          
          ask = MsgBox("Is your client uses laptop", vbYesNo + vbInformation)
          
          If ask = vbYes Then
               lvClients.ListItems(ItmDetails.ItemIndex).SubItems(2) = Time
               lvClients.ListItems(ItmDetails.ItemIndex).Tag = 1
          End If
          
          frmUserName.Show vbModal, Me
          frmServices.Show vbModal, Me
          DoEvents
          
          If sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected Then
               'sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData "UNAME|" & lvClients.ListItems(ItmDetails.ItemIndex).SubItems(1): DoEvents
               sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_LOGIN: DoEvents
          End If
          
          lvClients.ListItems(ItmDetails.ItemIndex).SubItems(3) = vbNullString
          
          For i = 1 To lvServList.ListItems.Count
               If lvServList.ListItems(i).Checked = True Then
                    Services = Services + lvServList.ListItems(i).Key + "=" + lvServList.ListItems(i).SubItems(3) + "|"
               End If
          Next i
          
          Services = Left$(Services, Len(Services) - 1)     ' remove last pipe
          
          Call SetTemporaryRecord(ItmDetails.ItemIndex, True, Services)
          
          'If sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected Then
          '     sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_LOGIN & "|" & lvClients.ListItems(ItmDetails.ItemIndex).SubItems(2): DoEvents
          
          'End If
          
          If ask = vbYes Then ItmDetails.IsLogedIn = True
          
          Toolbar.Buttons(1).Enabled = False
          For i = 2 To Toolbar.Buttons.Count
               Toolbar.Buttons(i).Enabled = True
          Next i
     Else
          MsgBox "Workstation is in use.", vbInformation
     End If
End Sub

Sub LogOut()
     Dim Items           As ListItem
     Dim TymUs           As Variant
     Dim ttl             As Currency
     Dim InetAmnt        As Currency
     Dim InetPM          As Currency
     
     Dim i               As Integer
          
     If ItmDetails.IsLogedIn Then
          lvClients.ListItems(ItmDetails.ItemIndex).Tag = 0
          lvServList.Enabled = False
          
          ' show logout time
          lvClients.ListItems(ItmDetails.ItemIndex).SubItems(3) = Time
          
          ' Compute Services
          ' ---------------------------------------------------------------------------------------------------------
          For i = 1 To lvServList.ListItems.Count
               If lvServList.ListItems(i).Checked = True Then
                    Set Items = frmPayment.lvSelService.ListItems.Add(, , lvServList.ListItems(i).Text)
                    Items.Key = lvServList.ListItems(i).Key
                    Items.SubItems(1) = lvServList.ListItems(i).SubItems(1)     ' Service Amount
                    Items.SubItems(2) = lvServList.ListItems(i).SubItems(2)     ' Unit
                    Items.SubItems(3) = lvServList.ListItems(i).SubItems(3)     ' Quantity
                    
                    If Items.SubItems(3) = vbNullString Then
                         Items.SubItems(3) = 0
                    End If
                    
                    Items.SubItems(4) = Items.SubItems(1) * Items.SubItems(3)   ' total amount
                    DoEvents
                    
                    ttl = ttl + Items.SubItems(4)
               End If
               
               lvServList.ListItems(i).Checked = False
          Next i
          
          ttl = FormatNumber(ttl, 2)
          frmPayment.txServAmnt.Text = "Php " & ttl
          frmPayment.txServAmnt.Tag = ttl
          
          Call AutoResizeListView(frmPayment.lvSelService)
          
          ttl = 0
          
          ' Compute Internet Usage
          ' ---------------------------------------------------------------------------------------------------------
          With frmPayment
               .lblUsrNme.Caption = lvClients.ListItems(ItmDetails.ItemIndex).SubItems(1)
               .lblUsrNme.Tag = lvClients.ListItems(ItmDetails.ItemIndex).Key
               .lblLogInfo(0) = Date
               .lblLogInfo(1) = lvClients.ListItems(ItmDetails.ItemIndex).SubItems(2)
               .lblLogInfo(2) = lvClients.ListItems(ItmDetails.ItemIndex).SubItems(3)
               TymUs = lvClients.ListItems(ItmDetails.ItemIndex).SubItems(4)
               .lblLogInfo(3) = TymUs
               
               ' get excess minute
               TymUs = Split(TymUs, ",")     ' 00h, 00m, 00s
               .lblLogInfo(4) = Trim$(TymUs(1))
               
               InetAmnt = frmPayment.lvSelService.ListItems(1).SubItems(1)
               InetPM = InetAmnt / 60
               
               ttl = InetAmnt * CInt(Left$(Trim$(TymUs(0)), 2))
               ttl = ttl + (InetPM * CInt(Left$(Trim$(TymUs(1)), 2)))
               ttl = FormatNumber(ttl, 2)
               
               .txLogAmnt.Text = "Php " & ttl
               .txLogAmnt.Tag = ttl
               
               ttl = CDbl(.txServAmnt.Tag) + CDbl(.txLogAmnt.Tag)
               ttl = FormatCurrency(ttl, 2)
               
               .txTtlAmnt.Text = "Php " & ttl
          End With
          
          frmPayment.Show vbModal, Me
          DoEvents
          
          ' Clear Selected Row
          ' ---------------------------------------------------------------------------------------------------------
          For i = 1 To lvClients.ColumnHeaders.Count - 1
               lvClients.ListItems(ItmDetails.ItemIndex).SubItems(i) = vbNullString
          Next i
          DoEvents
               
          ' send client CONN_LOGOUT message
          ' ---------------------------------------------------------------------------------------------------------
          If sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected Then
               sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_LOGOUT: DoEvents
          End If
          
          ' remove temporary record to database
          ' ---------------------------------------------------------------------------------------------------------
          Call RemoveTemporaryRecordBy(ItmDetails.IPNumber)
          
          ttl = 0
                    
          With ItmDetails
               .ItemIndex = 0
               .ItemName = vbNullString
               .ClientPCID = vbNullString
               .IPNumber = vbNullString
               .IsConnected = IIf(sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected, True, False)
               .IsLogedIn = False
          End With
     End If
End Sub

Sub TimeLimit(ByVal Button As MSComctlLib.Button)
     Dim temp       As String
     Dim LgTym      As String
               
     temp = Format(lvClients.ListItems(ItmDetails.ItemIndex).SubItems(3), "hh:mm:ss")
     LgTym = Format(lvClients.ListItems(ItmDetails.ItemIndex).SubItems(2), "hh:mm:ss")
                        
     ' is selected item is in use?
     If ItmDetails.IsLogedIn Then
          If temp <> vbNullString Then                 ' if Used Time Column is not empty
               txTL(0).Text = Hour(temp) - Hour(LgTym)           ' get hour and minute
               txTL(1).Text = Minute(temp) - Minute(LgTym)
          Else                                         ' else.. set 2 textbox to "00" value
               txTL(0).Text = "00"
               txTL(1).Text = "00"
          End If
     
          ihFloatTLFrame.Move Button.Left, Toolbar.Top + Toolbar.Height
          ihFloatTLFrame.Visible = True
          Button.Value = tbrPressed
          
          Call DrawShadow_IC(Me, WSM_Frame, ihFloatTLFrame, picFloatShadow(1), 2)
     End If
End Sub

Sub Change_Workstation(ByVal Button As MSComctlLib.Button)
     Dim i          As Integer
     
     If ItmDetails.IsLogedIn Then
          txSelWS.Text = ItmDetails.ItemName
          lsMovWS.Clear
          
          For i = 1 To lvClients.ListItems.Count
               'If (lvClients.ListItems(i).Tag = 0) And (lvClients.ListItems(i).SmallIcon = 1) Then
               If (lvClients.ListItems(i).Tag = 0) Then
                    lsMovWS.AddItem lvClients.ListItems(i).Text
               End If
          Next i
          
          ' show current user
          lblCW(0).Caption = lvClients.ListItems(ItmDetails.ItemIndex).SubItems(1)
          
          Button.Value = tbrPressed
          
          ' show floating menu
          ihFloatCngWS.Move Button.Left, Toolbar.Top + Toolbar.Height
          ihFloatCngWS.Visible = True: DoEvents
          Call DrawShadow_IC(Me, WSM_Frame, ihFloatCngWS, picFloatShadow(1), 2)
          ' ------------------
     End If
End Sub

Sub CancelUsage()
     Dim c          As Integer
     
     For c = 1 To lvClients.ColumnHeaders.Count - 1
          lvClients.ListItems(ItmDetails.ItemIndex).SubItems(c) = vbNullString
     Next c
     
     If sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected Then
          sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_CANCEL: DoEvents
     End If
     
     lvClients.ListItems(ItmDetails.ItemIndex).Tag = 0
     
     Call RemoveTemporaryRecordBy(ItmDetails.IPNumber)
     
     With ItmDetails
          .ItemIndex = 0
          .ItemName = vbNullString
          .ClientPCID = vbNullString
          .IPNumber = vbNullString
          .IsConnected = IIf(sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).State = sckConnected, True, False)
          .IsLogedIn = False
     End With
End Sub

' information
' -------------------------------------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------------------------------------
Private Sub Timer1_Timer()
     ShowInformation "Just hover your mouse pointer on Buttons, List, Toolbar for information."
End Sub

Private Sub lvServList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     ShowInformation lvServList.ToolTipText
End Sub

Private Sub lvClients_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     ShowInformation lvClients.ToolTipText
End Sub

Private Sub btnMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     ShowInformation btnMenu(Index).ToolTipText
End Sub

Private Sub Srv_ALbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     ShowInformation Srv_ALbl(Index).ToolTipText
End Sub

' Network Stuff
' -------------------------------------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------------------------------------
Private Sub Listen()
     sckServer(0).LocalPort = 2270
     sckServer(0).Listen
     
     ReDim Preserve WSAddVars(lvClients.ListItems.Count)
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
     Dim i          As Integer
     Dim res        As Integer
     
     SocketCount = SocketCount + 1
     Load sckServer(SocketCount)
               
     sckServer(SocketCount).Accept requestID
     
     res = IsDaClientComputerAllowed(lvClients, sckServer(SocketCount).RemoteHostIP)
     
     If res <> -1 Then
          WSAddVars(res).WinsockIndex = SocketCount
          
          SendData sckServer(SocketCount), CONN_CONNECTED: DoEvents
          
          Call SetConnectComputerEffect(lvClients, sckServer(SocketCount).RemoteHostIP, True)
     Else ' plan for attack
          sckServer(SocketCount).SendData "GotHUB? Internet - Server Side Application v1.0" + vbCrLf + _
                                          "You are illegally entering in our Server Side Application." + vbCrLf + _
                                          "Your connection will automatically close."
                                          DoEvents
                                          
          sckServer(SocketCount).Close: DoEvents
     End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
     On Error GoTo hell
     
     Dim sData      As String
     Dim msg()      As String
     Dim i          As Integer
     Dim temp()     As String
     Dim EnumWin()  As String
     Dim Items      As ListItem
     
     Static FylLen  As Long
     Dim perc       As Integer
     Dim FylDat     As String
     
     sckServer(Index).GetData sData$, vbString
     msg = Split(sData, "=")
     
     'MsgBox sData
     
     If sData$ = CONN_DISCONNECT Then
          SendData sckServer(Index), CONN_DISCONNECTED
               
          sckServer(Index).Close
          
          Call SetConnectComputerEffect(lvClients, sckServer(Index).RemoteHostIP, False)
          
     ElseIf msg(0) = CONN_SENDENUMWIN Then
          With frmMonitor
               .lvEnumList.ListItems.Clear
               
               temp = Split(msg(1), "|")
               
               For i = 0 To UBound(temp) - 1
                    EnumWin = Split(temp(i), ",")
                    
                    Set Items = .lvEnumList.ListItems.Add(, , EnumWin(1))
                    Items.Tag = EnumWin(0)
               Next i
               
               Call AutoResizeListView(.lvEnumList)
          End With
     ElseIf sData$ = CONN_REQUESTSTATUS Then
          i = GetClientStatus(lvClients, i, sckServer(Index).RemoteHostIP)
          DoEvents

          If i <> 0 Then
               INETRent = GetServiceAmount_IP(sckServer(Index).RemoteHostIP)
               
               sckServer(WSAddVars(i).WinsockIndex).SendData CONN_REQUESTSTATUS & "|" & _
                                                             lvClients.ListItems(i).SubItems(1) & "|" & _
                                                             lvClients.ListItems(i).SubItems(2) & "|" & _
                                                             Left$(lvClients.ListItems(i).SubItems(4), 8) & "|" & _
                                                             INETRent
               DoEvents
          End If
     ElseIf msg(0) = CONN_FILESTAT Then
          If msg(1) = "1" Then
               Open tempScrFyl For Binary As #1
               DoEvents
               
               ttlsnt = 0
               
               sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData CONN_FILESTAT & "|" & "2"
               DoEvents
               
               FylLen = msg(2)
          ElseIf msg(1) = "3" Then
               FylDat$ = Mid$(sData$, 6, Len(sData$))
               
               Put #1, , FylDat$
               DoEvents
               
               ttlsnt = ttlsnt + Len(FylDat$)
               
               perc = CInt(ttlsnt / FylLen * 100)
               
               frmMonitor.ProgressBar.Value = perc
               frmMonitor.lblRef.Caption = "Refreshing... " & perc & "%"
               DoEvents
               
               Sleep 200
          ElseIf msg(1) = "4" Then
               Close #1
               DoEvents
               
               frmMonitor.imgScr.Picture = LoadPicture(tempScrFyl)
               
               Kill tempScrFyl
               DoEvents
               
               If frmMonitor.ctlLbl(3).Enabled = False Then frmMonitor.ctlLbl(3).Enabled = True
          End If
     ElseIf sData = CONN_LOGEDIN Then
          lvClients.ListItems(ItmDetails.ItemIndex).SubItems(2) = Time
          lvClients.ListItems(ItmDetails.ItemIndex).Tag = 1
          
          sckServer(WSAddVars(ItmDetails.ItemIndex).WinsockIndex).SendData "UNAME|" & _
                                                                           lvClients.ListItems(ItmDetails.ItemIndex).SubItems(1) & "|" & _
                                                                           INETRent: DoEvents
     ElseIf sData = CONN_LOGOUT Then
          
     End If
     
     Exit Sub
hell:
     If Err.Number = 481 Then
          ' notification error when invalid screenshot file was send
          MsgBox "Unable to refresh the screen. Please click the Refresh Button to try again.", vbExclamation, "Screen Capture"
     Else
          MsgBox Err.Description + vbCrLf & Err.Number, vbCritical, "error in function sckServeR_DataArrival"
     End If
     
     If frmMonitor.ctlLbl(3).Enabled = False Then frmMonitor.ctlLbl(3).Enabled = True
End Sub

' Resizing stuff
' ------------------------------------------------------------------------------------------------------------------
Private Sub Form_Resize()
     On Error Resume Next          ' error occured when window is minimize. so just resume anyway
     
     'If m_Sizing Then Exit Sub
     
     WSM_Frame.Move Frame(0).Left + Frame(0).Width + tX(4), TopFrame.Top + TopFrame.Height + tY(8)
     WSM_Frame.Width = ScaleWidth - WSM_Frame.Left - tX(4)
     
     Frame(1).Move WSM_Frame.Left, ScaleHeight - Frame(1).Height - StatusBar.Height
     Frame(1).Width = WSM_Frame.Width
     Frame(1).Height = ScaleHeight - StatusBar.Height - Frame(1).Top - tY(4)
     
     'WSM_Frame.Height = ScaleHeight - StatusBar.Height - WSM_Frame.Top - tY(6)
     WSM_Frame.Height = ScaleHeight - WSM_Frame.Top - (StatusBar.Height + tY(8) + Frame(1).Height) - tY(4)
End Sub

Private Sub picMenuFrame_Resize()
     'If m_Sizing Then Exit Sub
     
     'Call DrawShadow(Me, TopFrame, picOist, 3)
     'Call Gradient(picMenuFrame, RGB(199, 211, 247), RGB(51, 102, 204), "VERTICAL")
End Sub

Private Sub Frame_Resiz_e(Index As Integer)
     Dim i          As Integer
     
     'If m_Sizing Then Exit Sub
     
     lvbtns(0).Move Frame(1).ScaleWidth - lvbtns(0).Width - 5, (Frame(1).ScaleHeight - lvbtns(0).Height) / 2
     
     For i = 1 To lvbtns.Count - 1
          lvbtns(i).Move lvbtns(i - 1).Left - lvbtns(i).Width - 4, lvbtns(0).Top
     Next i
     
     'IH(4).Top = Frame(0).ScaleHeight - IH(4).Height - tY(5)
     'Call DrawIHPanelsShadow(2)
     'Call Gradient(Frame(0), RGB(122, 161, 230), vbWhite, "VERTICAL")
End Sub
