VERSION 5.00
Begin VB.Form frmANS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Service"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
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
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btns 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   345
      Index           =   1
      Left            =   2070
      TabIndex        =   8
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton btns 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   3375
      TabIndex        =   7
      Top             =   1470
      Width           =   1215
   End
   Begin VB.TextBox txSI 
      Height          =   285
      Index           =   2
      Left            =   2595
      TabIndex        =   6
      Top             =   915
      Width           =   1905
   End
   Begin VB.TextBox txSI 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1005
      TabIndex        =   4
      Top             =   915
      Width           =   810
   End
   Begin VB.TextBox txSI 
      Height          =   285
      Index           =   0
      Left            =   1005
      TabIndex        =   2
      Top             =   570
      Width           =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   4680
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   4680
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Per Unit"
      Height          =   195
      Index           =   2
      Left            =   1950
      TabIndex        =   5
      Top             =   960
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   195
      Index           =   1
      Left            =   585
      TabIndex        =   3
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   525
      TabIndex        =   1
      Top             =   615
      Width           =   405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   435
      X2              =   4590
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   105
      Picture         =   "frmANS.frx":0000
      Top             =   105
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Service Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   450
      TabIndex        =   0
      Top             =   150
      Width           =   2085
   End
End
Attribute VB_Name = "frmANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btns_Click(Index As Integer)
     Dim res        As Boolean
     
     If Index = 0 Then
          If IsValid Then
               With SelService
                    .Name = txSI(0).Text
                    .Rate = txSI(1).Text
                    .Unit = txSI(2).Text
               End With
               
               If Tag = "ADD" Then
                    res = AddNewService
                    DoEvents
               ElseIf Tag = "EDIT" Then
                    res = EditService
                    DoEvents
               End If
               
               If res = True Then
                    Call GetServices(frmServMngr.lvServices)
                    DoEvents
               End If
               
               Unload Me
          End If
     ElseIf Index = 1 Then
          Unload Me
     End If
End Sub

Private Sub txSI_GotFocus(Index As Integer)
     txSI(Index).SelStart = 0
     txSI(Index).SelLength = Len(txSI(Index).Text)
End Sub

Function IsValid() As Boolean
     Dim Cntrls          As Control
     Dim i               As Integer
     
     IsValid = False
     
     For Each Cntrls In Controls
          If TypeOf Cntrls Is TextBox Then
               If Cntrls.Text <> vbNullString Then
                    i = i + 1
               Else
                    MsgBox "You must enter the '" + Label2(Cntrls.Index).Caption + "'", vbInformation
               End If
          End If
     Next
     
     If i = 3 Then
          IsValid = True
          
          If Not IsNumeric(txSI(1).Text) Then
               MsgBox "'Rate' must be a numerical value.", vbInformation
               
               IsValid = False
          End If
     Else
          IsValid = False
     End If
End Function
