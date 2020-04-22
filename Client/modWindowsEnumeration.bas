Attribute VB_Name = "modWindowsEnumeration"
Option Explicit

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Integer) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const GW_OWNER = 4
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_APPWINDOW = &H40000
Private Const WM_GETICON = &H7F
Private Const GCL_HICON = (-14)
Private Const GCL_HICONSM = (-34)
Private Const WM_QUERYDRAGICON = &H37

Public sEnumed           As String

Function EnumWnd(ByVal hwnd As Long, ByVal lParam As Long) As Long
     Dim s_Caption       As String * 512
     Dim l_CapLen        As Long
     Dim ClassName       As String * 50
     
     l_CapLen = GetWindowTextLength(hwnd)
     GetWindowText hwnd, s_Caption, l_CapLen + 1
     GetClassName hwnd, ClassName, 50

     Call Insert(hwnd, s_Caption)
     
     EnumWnd = 1
End Function

Private Sub Insert(ByVal hwnd As Long, ByVal WndCaption As String)
     Dim i               As Long
     Dim l_icon          As Long
     Dim cnt             As Long
     Dim HasNoOwner      As Long
     Dim WindowStyle     As Long
     
     If IsWindowVisible(hwnd) Then
          If GetParent(hwnd) = 0 Then
               HasNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
               WindowStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
               If (((WindowStyle And WS_EX_TOOLWINDOW) = 0) And HasNoOwner) Or ((WindowStyle And WS_EX_APPWINDOW) And Not HasNoOwner) Then
'                    With frmMonitor
'                         l_icon = GetIcon(hwnd)
'
'                         .Ico.Cls
'
'                         DrawIconEx .Ico.hdc, 0, 0, l_icon, 16, 16, 0, 0, 3
'
'                         .ImageList1.ListImages.Add , , .Ico.Image
'                         cnt = .ImageList1.ListImages.Count
'
'                         .Ico.Refresh
'
'                         lv.ListItems.Add , , WndCaption, , .ImageList1.ListImages(cnt).Index
'                         i = lv.ListItems.Count
'                    End With

                    sEnumed = sEnumed + CStr(hwnd) + "," + WndCaption + "|"
               End If
          End If
     End If
End Sub

Public Function GetIcon(hwnd As Long) As Long
    Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICONSM)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICON)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

