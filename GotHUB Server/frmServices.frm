VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServices 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Services"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6600
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btn 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   5325
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvServices 
      Height          =   2985
      Left            =   60
      TabIndex        =   1
      Top             =   675
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5265
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Services"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   585
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1032
      InfoHeaderStyle =   0
      HeaderHeight    =   26
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
      Caption         =   $"frmServices.frx":0000
      MultiLine       =   -1  'True
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmServices.frx":0043
      EdgeRadiusSize  =   12
      RoundEdgeStyle  =   6
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Click()
     frmMain.lvServList.Enabled = True
     Unload Me
End Sub

Private Sub Form_Load()
     IH(0).EdgeRadiusSize = 12
     IH(0).LeftColor = RGB(193, 210, 237)
     IH(0).RightColor = vbWhite
     IH(0).ForeColor = vbBlack
     
     Call StyleStaticEdge(lvServices)
     
     Call GetServices(lvServices)
     DoEvents
     
     lvServices.ListItems(1).Checked = True
     INETRent = lvServices.ListItems(1).SubItems(1)
     frmMain.lvServList.ListItems(1).Checked = True
End Sub

Private Sub lvServices_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Dim Items       As ListItem
     Dim askQuan     As String
     
     If Item.Index <> 1 And Item.Index <> 2 Then
          askQuan = 0
          askQuan = InputBox("Enter quantity for selected service", , "1")
          
          If Not IsNumeric(askQuan) Then
               MsgBox "You must enter an numerical value", vbInformation
               
               Call lvServices_ItemCheck(Item)
               
               Exit Sub
          End If
          
          If askQuan = vbNullString Or askQuan = 0 Then askQuan = 0: Item.Checked = False
          
          lvServices.ListItems(Item.Index).SubItems(3) = askQuan
          frmMain.lvServList.ListItems(Item.Index).SubItems(3) = askQuan
     End If
     
     ' we dont want to have a both check marks on Internet Service and Internet with WebCam Service.
     ' so heres how to prevent that
     ' --------------------------------------------------------------------------------------------------------------
     If Item.Index = 1 Or Item.Index = 2 Then
          lvServices.ListItems(1).Checked = False
          lvServices.ListItems(2).Checked = False
          frmMain.lvServList.ListItems(1).Checked = False
          frmMain.lvServList.ListItems(2).Checked = False
     
          Item.Checked = True
          frmMain.lvServList.ListItems(Item.Index).Checked = Item.Checked
          
          INETRent = Item.SubItems(1)
     End If
     ' --------------------------------------------------------------------------------------------------------------
     
     Set Items = frmMain.lvServList.ListItems(Item.Index)
     Items.Checked = Item.Checked
     Items.Selected = True
     Items.EnsureVisible
End Sub
