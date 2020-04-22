VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWalkthroughServ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Walkthrough Service"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWalkthroughServ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btns 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   1
      Left            =   4695
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton btns 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   5985
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txNWTC 
      Height          =   285
      Left            =   2175
      TabIndex        =   0
      Top             =   570
      Width           =   4950
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   90
      Width           =   7125
      _ExtentX        =   12568
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
      Caption         =   "Walkthrough Service"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmWalkthroughServ.frx":058A
      EdgeRadiusSize  =   12
      RoundEdgeStyle  =   0
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   4710
      Index           =   1
      Left            =   75
      TabIndex        =   3
      Top             =   990
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   8308
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
      Caption         =   "Services List - mark the service you want to enter quantity."
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   0   'False
      Picture         =   "frmWalkthroughServ.frx":0B24
      EdgeRadiusSize  =   0
      RoundEdgeStyle  =   7
      Begin MSComctlLib.ListView lvServices 
         Height          =   4140
         Left            =   120
         TabIndex        =   4
         Top             =   450
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   7303
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      Caption         =   "Name of walkthrough client"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   615
      Width           =   1950
   End
End
Attribute VB_Name = "frmWalkthroughServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub InitColumns()
     With lvServices
          .ColumnHeaders.Add , , "Service Name"
          .ColumnHeaders.Add , , "Rate", , lvwColumnRight
          .ColumnHeaders.Add , , "Per Unit"
          .ColumnHeaders.Add , , "Quantity"
          
          .LabelEdit = lvwAutomatic
     End With
End Sub

Private Sub btns_Click(Index As Integer)
     Dim i          As Integer
     Dim cnt        As Integer
     Dim ttl        As Currency
     
     Dim res        As Boolean
     
     Dim Total      As Currency
     
     If Index = 0 Then
          If Valid Then
               With FullRec
                    .Name = txNWTC.Text
                    
                    For i = 1 To lvServices.ListItems.Count
                         If (lvServices.ListItems(i).Checked = True) And (lvServices.ListItems(i).SubItems(3) <> 0) Then
                              cnt = cnt + 1
                              
                              ReDim Preserve .ServiceID(cnt - 1)
                              ReDim Preserve .QUantity(cnt - 1)
                              ReDim Preserve .Amount(cnt - 1)
                              
                              ttl = CCur(lvServices.ListItems(i).SubItems(3) * _
                                         lvServices.ListItems(i).SubItems(1))
                              
                              .ServiceID(cnt - 1) = lvServices.ListItems(i).Key
                              .QUantity(cnt - 1) = lvServices.ListItems(i).SubItems(3)
                              .Amount(cnt - 1) = ttl
                              
                              Total = Total + ttl
                         End If
                    Next i
                    
                    .LogOutTime = Time
               End With
               
               res = CreateWalkthroughRecord
               
               MsgBox "Total service amount: Php" & FormatNumber(Total, 2), vbInformation, "Walkthrough Service"
               
               If res = True Then
                    MsgBox "Walkthrough service saved.", vbInformation, "Walkthrough Service"
                    
                    Unload Me
               End If
          End If
     ElseIf Index = 1 Then
          Unload Me
     End If
End Sub

Private Sub Form_Load()
     Call InitColumns
     
     Call GetServices(lvServices)
     Call AutoResizeListView(lvServices)
     Call StyleStaticEdge(lvServices)
     
     ' its just a walktrhrough service
     ' so just remove the the Internet service
     lvServices.ListItems.Remove 2 ' remove Internet
     lvServices.ListItems.Remove 1 ' remove Internet /w CAM
     
     btns(0).Enabled = True
End Sub

Private Sub lvServices_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Dim ask        As String
     Dim def        As String
     
     def = IIf(lvServices.ListItems(Item.Index).SubItems(3) <> 0, _
               lvServices.ListItems(Item.Index).SubItems(3), "1")
     
     If def <> 0 Then Item.Checked = True
     
     ask = InputBox("Enter service quantity", "Quantity", def)
     If ask = vbNullString Then Item.Checked = False
     
     If IsNumeric(ask) Then
          If ask = 0 Then Item.Checked = False
          
          lvServices.ListItems(Item.Index).SubItems(3) = ask
     Else
          Item.Checked = False
          
          Exit Sub
     End If
End Sub

Function Valid() As Boolean
     Dim i          As Integer
     
     If txNWTC.Text <> vbNullString Then
          Valid = False
          
          For i = 1 To lvServices.ListItems.Count
               If (lvServices.ListItems(i).Checked = True) And (lvServices.ListItems(i).SubItems(3) <> 0) Then
                    Valid = True
               End If
          Next i
          
          If Valid = False Then
               MsgBox "No services selected.", vbInformation
          End If
     Else
          MsgBox "You must enter the clients name.", vbInformation
          Valid = False
     End If
End Function
