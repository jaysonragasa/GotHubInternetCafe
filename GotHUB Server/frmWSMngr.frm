VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmWSMngr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Workstation Manager"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWSMngr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   5985
   End
   Begin VB.CommandButton btns 
      Caption         =   "&Close"
      Height          =   345
      Index           =   1
      Left            =   6270
      TabIndex        =   8
      Top             =   6045
      Width           =   1485
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   4965
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   8758
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
      Caption         =   "Workstation List"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   0   'False
      Picture         =   "frmWSMngr.frx":058A
      EdgeRadiusSize  =   12
      RoundEdgeStyle  =   6
      Begin MSMask.MaskEdBox txIPAdd 
         Height          =   285
         Left            =   2085
         TabIndex        =   10
         Top             =   4575
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   15
         Format          =   "###.###.###.###"
         Mask            =   "###.###.###.###"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton btns 
         Caption         =   "&Update Settings"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4035
         TabIndex        =   9
         Top             =   4575
         Width           =   1485
      End
      Begin VB.TextBox txWMInfo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   4575
         Width           =   1875
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8160
         Top             =   3090
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWSMngr.frx":0B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWSMngr.frx":13FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWSMngr.frx":1CD8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvWStxn 
         Height          =   3825
         Left            =   105
         TabIndex        =   4
         Top             =   420
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   6747
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   7
         Top             =   4335
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Workstation Name"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   4335
         Width           =   1320
      End
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   688
      InfoHeaderStyle =   0
      HeaderHeight    =   26
      FrameBackColor  =   16777215
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
      Caption         =   "Workstation Manager"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmWSMngr.frx":29B2
      EdgeRadiusSize  =   12
      RoundEdgeStyle  =   4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage your workstations. You can change the IP Address, Workstation Name settings."
      Height          =   390
      Index           =   1
      Left            =   1245
      TabIndex        =   2
      Top             =   495
      Width           =   7815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   495
      Width           =   1005
   End
End
Attribute VB_Name = "frmWSMngr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelIP      As String
Dim SelIndex   As Integer
Dim NowUpdate  As Boolean
Dim temp()     As String

Private Sub btns_Click(Index As Integer)
     Dim res        As Boolean
     
     If Index = 0 Then
          res = Update_WS_IPNumber(SelIP, txWMInfo(0).Text, txIPAdd.Text)
          Call UpdateLVI_WS_IPNumber(frmMain.lvClients, SelIndex, txWMInfo(0).Text, txIPAdd.Text)
          Call UpdateLVI_WS_IPNumber(lvWStxn, SelIndex, txWMInfo(0).Text, txIPAdd.Text)
          
          If res = True Then
               MsgBox "workstation settings updated.", vbInformation
                      
               Timer1.Enabled = True
          End If
     ElseIf Index = 1 Then
          Unload Me
     End If
End Sub

Sub UpdateLVI_WS_IPNumber(ByRef ctlListView As ListView, ByVal SelIndex As Integer, ByVal NewWSName As String, ByVal NewIPN As String)
     ctlListView.ListItems(SelIndex).Text = NewWSName
     ctlListView.ListItems(SelIndex).Key = temp(0) + "|" + NewIPN
End Sub

Private Sub btnTest_Click()
     If SelIndex <> 0 Then
     End If
End Sub

Private Sub Form_Load()
     Call GetWorkstations
     
     IH(0).LeftColor = SystemColorConstants.vbActiveTitleBar
     IH(0).RightColor = SystemColorConstants.vb3DFace
     IH(0).ForeColor = SystemColorConstants.vbActiveTitleBarText
     
     IH(1).LeftColor = SystemColorConstants.vb3DFace
     IH(1).RightColor = SystemColorConstants.vbWindowBackground
     IH(1).ForeColor = SystemColorConstants.vbButtonText
     
     IsWMLoaded = True
End Sub

Private Sub lvWStxn_ItemClick(ByVal Item As MSComctlLib.ListItem)
     Dim tmp()      As String
     Dim IP         As String
     
     temp = Split(Item.Key, "|")
     
     tmp = Split(temp(1), ".")
     
     IP = Right$("000", 3 - Len(tmp(0))) & tmp(0) & "."
     IP = IP + Right$("000", 3 - Len(tmp(1))) & tmp(1) & "."
     IP = IP + Right$("000", 3 - Len(tmp(2))) & tmp(2) & "."
     IP = IP + Right$("000", 3 - Len(tmp(3))) & tmp(3)
     
     txWMInfo(0).Enabled = IIf(Item.Icon = 3 Or Item.Icon = 1, False, True)
     txIPAdd.Enabled = IIf(Item.Icon = 3 Or Item.Icon = 1, False, True)
     btns(0).Enabled = IIf(Item.Icon = 3 Or Item.Icon = 1, False, True)
     
     txWMInfo(0).Text = Item.Text
     txIPAdd.Text = IP
     
     SelIP = IIf(Item.Icon = 3 Or Item.Icon = 1, vbNullString, temp(1))
     SelIndex = IIf(Item.Icon = 3 Or Item.Icon = 1, 0, Item.Index)
End Sub

Private Sub Timer1_Timer()
     Dim Items           As ListItem
     lvWStxn.ListItems.Clear
     
     Call GetWorkstations
     
     Set Items = lvWStxn.ListItems(SelIndex)
     Items.Selected = True
     Items.EnsureVisible
     lvWStxn.SetFocus
     
     txIPAdd.SetFocus
End Sub

Private Sub txWMInfo_GotFocus(Index As Integer)
     AutoHighlight txWMInfo(Index)
End Sub

Sub GetWorkstations()
     Dim i          As Integer
     
     For i = 1 To frmMain.lvClients.ListItems.Count
          If frmMain.lvClients.ListItems(i).SmallIcon = 2 Then
               lvWStxn.ListItems.Add , frmMain.lvClients.ListItems(i).Key, frmMain.lvClients.ListItems(i).Text, 2
          Else
               If frmMain.lvClients.ListItems(i).SubItems(1) = vbNullString Then
                    lvWStxn.ListItems.Add , frmMain.lvClients.ListItems(i).Key, frmMain.lvClients.ListItems(i).Text, 1
               Else
                    lvWStxn.ListItems.Add , frmMain.lvClients.ListItems(i).Key, "Workstation Inuse", 3
               End If
          End If
     Next i
End Sub
