Attribute VB_Name = "modAPI"
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

Private Const LVM_FIRST = &H1000
Public Const LVM_SETTEXTBKCOLOR As Long = (LVM_FIRST + 38)
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

Private Const LVM_GETHEADER = (&H1000 + 31)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const HDS_BUTTONS = &H2
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_EX_DLGMODALFRAME = &H1

Public Function Flatten_ListView_ColumnButton(hwnd As ListView)
    Dim lS          As Long
    Dim lHwnd       As Long
    Dim wHwnd       As Long
    
    wHwnd = hwnd.hwnd
    lHwnd = SendMessageByLong(wHwnd, LVM_GETHEADER, 0, 0)
    
    If (lHwnd <> 0) Then
        lS = GetWindowLong(lHwnd, GWL_STYLE)
        lS = lS And Not HDS_BUTTONS
        SetWindowLong lHwnd, GWL_STYLE, lS
    End If
End Function

Sub AutoResizeListView(ByRef ctlListView As ListView)
     Dim Column          As Long
     Dim Counter         As Long
     
     Counter = 0
     
     For Column = Counter To ctlListView.ColumnHeaders.Count - 1
          SendMessage ctlListView.hwnd, LVM_SETCOLUMNWIDTH, Column, LVSCW_AUTOSIZE_USEHEADER
     Next
End Sub

Sub StyleStaticEdge(ByRef ctlObject As Object)
     Dim ret             As Long
     
     ret = SetWindowLong(ctlObject.hwnd, GWL_EXSTYLE, WS_EX_STATICEDGE)
     ctlObject.Width = ctlObject.Width + tX(1)
     ctlObject.Width = ctlObject.Width - tX(1)
End Sub

Sub DrawModalFrame(ByVal ohWnd As Long)
     SetWindowLong ohWnd, GWL_EXSTYLE, WS_EX_DLGMODALFRAME
End Sub

Public Sub sysControlHover(ctl As Control, x As Single, y As Single)
     Dim HitTest As Long
     
     On Error Resume Next
     
     'test for control hwnd, if error simply exit
     
     HitTest = ctl.hwnd
     
     If Err.Number <> 0 Then Exit Sub
     
     With ctl
          If (x < 0) Or (y < 0) Or (x > .Width) Or (y > .Height) Then
               ReleaseCapture
               .Visible = False
          Else
               SetCapture .hwnd
               .Visible = True
          End If
     End With
     
     On Error GoTo 0
End Sub

Public Sub DisableXButton(vForm As Form)
    ' Sub/Function Name       : frmKillExit
    ' Purpose                 : Disables the X button on a form
    ' Parameters              : Form Object

    Dim hSysMenu As Long
    Dim nCnt As Long
    ' Get handle to our form's system menu
    ' (Restore, Maximize, Move, close etc.)
    hSysMenu = GetSystemMenu(vForm.hwnd, False)
    
    If hSysMenu Then
        ' Get System menu's menu count
        nCnt = GetMenuItemCount(hSysMenu)
        
        If nCnt Then
            ' Menu count is based on 0 (0, 1, 2, 3...)
            RemoveMenu hSysMenu, nCnt - 3, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 4, MF_BYPOSITION Or MF_REMOVE ' Remove the seperator
            'DrawMenuBar vForm.hwnd
            ' Force caption bar's refresh. Disabling X button
        End If
    End If
End Sub
