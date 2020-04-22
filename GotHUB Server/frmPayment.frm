VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPayment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Payment"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btns 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   3930
      TabIndex        =   19
      Top             =   7095
      Width           =   1215
   End
   Begin VB.TextBox txTtlAmnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Php0.00"
      Top             =   6435
      Width           =   4125
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   2250
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3975
      Width           =   5040
      _extentx        =   8890
      _extenty        =   3969
      infoheaderstyle =   1
      headerheight    =   21
      framebackcolor  =   16777215
      framebordercolor=   10070188
      gradientstyle   =   1
      leftcolor       =   14215660
      rightcolor      =   16777215
      margin          =   4
      fontname        =   "frmPayment.frx":0000
      forecolor       =   -2147483630
      caption         =   "Log Information"
      multiline       =   0   'False
      alignment       =   0
      hasicon         =   -1  'True
      picture         =   "frmPayment.frx":0028
      edgeradiussize  =   12
      roundedgestyle  =   1
      Begin VB.TextBox txLogAmnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Php0.00"
         Top             =   1830
         Width           =   1770
      End
      Begin VB.Image Image1 
         Height          =   15
         Index           =   2
         Left            =   60
         Picture         =   "frmPayment.frx":03C2
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   4890
      End
      Begin VB.Label lblLogInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hh:mm:ss"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3705
         TabIndex        =   16
         Top             =   855
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Execess Time:"
         Height          =   195
         Index           =   4
         Left            =   2610
         TabIndex        =   15
         Top             =   855
         Width           =   1020
      End
      Begin VB.Label lblLogInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hh:mm:ss"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3705
         TabIndex        =   14
         Top             =   420
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Used:"
         Height          =   195
         Index           =   3
         Left            =   2835
         TabIndex        =   13
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblLogInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hh:mm:ss am"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1275
         TabIndex        =   12
         Top             =   1305
         Width           =   1125
      End
      Begin VB.Label lblLogInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hh:mm:ss am"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   11
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label lblLogInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dd-mm-yy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1275
         TabIndex        =   10
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log out Time:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1305
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LogIn Time:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   8
         Top             =   855
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LogIn Date:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1950
         TabIndex        =   6
         Top             =   1890
         Width           =   1155
      End
   End
   Begin GoHUB_Server.InfoHeader IH 
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5040
      _extentx        =   8890
      _extenty        =   5741
      infoheaderstyle =   1
      headerheight    =   21
      framebackcolor  =   16777215
      framebordercolor=   10070188
      gradientstyle   =   1
      leftcolor       =   14215660
      rightcolor      =   16777215
      margin          =   4
      fontname        =   "frmPayment.frx":085C
      forecolor       =   -2147483630
      caption         =   "Selected Services"
      multiline       =   0   'False
      alignment       =   0
      hasicon         =   -1  'True
      picture         =   "frmPayment.frx":0884
      edgeradiussize  =   12
      roundedgestyle  =   1
      Begin VB.TextBox txServAmnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Php0.00"
         Top             =   2820
         Width           =   1770
      End
      Begin MSComctlLib.ListView lvSelService 
         Height          =   2385
         Left            =   75
         TabIndex        =   1
         Top             =   375
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   4207
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
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
            Text            =   "Per Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Total Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Service are not included"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   975
         TabIndex        =   20
         Top             =   3000
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount of Services"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   2805
         Width           =   2130
      End
   End
   Begin VB.Label lblUsrNme 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[name]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2175
      TabIndex        =   22
      Top             =   165
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   660
      TabIndex        =   21
      Top             =   165
      Width           =   1425
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   120
      Picture         =   "frmPayment.frx":0E1E
      Top             =   120
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   15
      Index           =   0
      Left            =   105
      Picture         =   "frmPayment.frx":17B8
      Stretch         =   -1  'True
      Top             =   6975
      Width           =   5085
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   6510
      Width           =   720
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     Dim i          As Integer
     
     For i = 0 To IH.Count - 1
          IH(i).LeftColor = SystemColorConstants.vb3DFace
          IH(i).RightColor = SystemColorConstants.vbWindowBackground
          IH(i).ForeColor = SystemColorConstants.vbButtonText
     Next i
     
     Call StyleStaticEdge(lvSelService)
     Call Flatten_ListView_ColumnButton(lvSelService)
End Sub

'Public Type FullRecord
'     Name                     As String      ' name of serviced person
'     IPNumber                 As String      ' ip address of the workstation used (if any)
'     LogInDate                As String
'     LogInTime                As String
'     LogOutTime               As String
'     TimeUsed                 As String
'     ServiceID()              As String
'     QUantity()               As Integer
'     Amount()                 As Currency
'End Type

Private Sub btns_Click()
     Dim i          As Integer
     Dim res        As Boolean
     Dim temp()     As String
     
     temp = Split(lblUsrNme.Tag, "|")
     With FullRec
          .Name = lblUsrNme.Caption
          .ClientPCID = temp(0)
          .IPNumber = temp(1)
          .LogInDate = Date
          .LogInTime = lblLogInfo(1).Caption
          .LogOutTime = lblLogInfo(2).Caption
          .TimeUsed = lblLogInfo(3).Caption
          .IU_Amount = txLogAmnt.Tag
          
          ReDim .ServiceID(lvSelService.ListItems.Count - 1)
          ReDim .QUantity(lvSelService.ListItems.Count - 1)
          ReDim .Amount(lvSelService.ListItems.Count - 1)
          
          For i = 1 To lvSelService.ListItems.Count
               .ServiceID(i - 1) = lvSelService.ListItems(i).Key
               .QUantity(i - 1) = lvSelService.ListItems(i).SubItems(3)
               .Amount(i - 1) = lvSelService.ListItems(i).SubItems(4)
          Next i
     End With
     
     res = CreateRecord
     
     If res = True Then
          MsgBox "All records of '" + lblUsrNme.Caption + "' is successfully saved in database", vbInformation
     End If
     
     Unload Me
End Sub

