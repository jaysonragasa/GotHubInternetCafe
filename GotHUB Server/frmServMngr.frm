VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E6E03A98-C7DC-4FCE-800D-724A332410A9}#1.0#0"; "LaVolpeButtons.ocx"
Begin VB.Form frmServMngr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Services Mangager"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServMngr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   6015
      TabIndex        =   8
      Top             =   6810
      Width           =   1215
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   7155
      _ExtentX        =   12621
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
      Caption         =   "Services Manganger"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmServMngr.frx":058A
      EdgeRadiusSize  =   12
      RoundEdgeStyle  =   4
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   5790
      Index           =   1
      Left            =   75
      TabIndex        =   3
      Top             =   945
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   10213
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
      Caption         =   "Services List"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   0   'False
      Picture         =   "frmServMngr.frx":0B24
      EdgeRadiusSize  =   0
      RoundEdgeStyle  =   7
      Begin LaVolpeButtons.lvButtons_H lvbtns 
         Height          =   360
         Index           =   2
         Left            =   3255
         TabIndex        =   7
         Top             =   435
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   635
         Caption         =   "Remove Service"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Image           =   "frmServMngr.frx":10BE
         cBack           =   -2147483633
      End
      Begin LaVolpeButtons.lvButtons_H lvbtns 
         Height          =   360
         Index           =   1
         Left            =   1530
         TabIndex        =   6
         Top             =   435
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   635
         Caption         =   "Add New Service"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Image           =   "frmServMngr.frx":1658
         cBack           =   -2147483633
      End
      Begin LaVolpeButtons.lvButtons_H lvbtns 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   435
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         Caption         =   "Edit Service"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmServMngr.frx":1BF2
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView lvServices 
         Height          =   4770
         Left            =   120
         TabIndex        =   4
         Top             =   885
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   8414
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
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
      TabIndex        =   2
      Top             =   495
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage your services. Modify your services rates, name, add new service, or remove."
      Height          =   450
      Index           =   1
      Left            =   1245
      TabIndex        =   1
      Top             =   495
      Width           =   5865
   End
End
Attribute VB_Name = "frmServMngr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
     Call GetServices(frmMain.lvServList)
     DoEvents
     
     Unload Me
End Sub

Private Sub Form_Load()
     
     Call InitColumns
     
     Call GetServices(lvServices)
     Call AutoResizeListView(lvServices)
     Call StyleStaticEdge(lvServices)
     
     IH(0).LeftColor = SystemColorConstants.vbActiveTitleBar
     IH(0).RightColor = SystemColorConstants.vb3DFace
     IH(0).ForeColor = SystemColorConstants.vbActiveTitleBarText
     
     IH(1).LeftColor = SystemColorConstants.vb3DFace
     IH(1).RightColor = SystemColorConstants.vbWindowBackground
     IH(1).ForeColor = SystemColorConstants.vbButtonText
     
     lvServices.Appearance = ccFlat
End Sub

Sub InitColumns()
     With lvServices
          .ColumnHeaders.Add , , "Service Name"
          .ColumnHeaders.Add , , "Rate", , lvwColumnRight
          .ColumnHeaders.Add , , "Per Unit"
          
          .LabelEdit = lvwAutomatic
     End With
End Sub

Private Sub lvbtns_Click(Index As Integer)
     If (SelService.ServiceID = vbNullString) And (Index <> 1) Then
          MsgBox "Please select any service from the list.", vbInformation
          
          Exit Sub
     End If
     
     If Index = 0 Then
          With frmANS
               .Caption = "Edit Selected Service"
               .Tag = "EDIT"
               
               .txSI(0).Text = SelService.Name
               .txSI(1).Text = SelService.Rate
               .txSI(2).Text = SelService.Unit
               
               .Show vbModal, Me
          End With
          
     ElseIf Index = 1 Then
          frmANS.Caption = "Add New Service"
          frmANS.Tag = "ADD"
          
          frmANS.Show vbModal, Me
          
     ElseIf Index = 2 Then
          Dim res As Boolean
          
          res = DeleteService
          
          If res = True Then
               With lvServices
                    .FindItem SelService.Name, , , lvwPartial
                    .ListItems.Remove .SelectedItem.Index
               End With
          End If
     End If
End Sub

Private Sub lvServices_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With SelService
          .ServiceID = Item.Key
          .Name = Item.Text
          .Rate = Item.SubItems(1)
          .Unit = Item.SubItems(2)
     End With
End Sub

