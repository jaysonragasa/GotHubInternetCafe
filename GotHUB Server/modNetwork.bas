Attribute VB_Name = "modNetwork"
Option Explicit

Public Type WSAdditionalVariables
     WinsockIndex                       As Integer
     ListViewIndex                      As Integer
End Type

Public IsConnected                      As Boolean
Public WSAddVars()                      As WSAdditionalVariables
Public SocketCount                      As Integer
Public BlckrCnxnCnt                     As Integer

Public IsWMLoaded                       As Boolean

Sub SendData(ByRef SockControl As Winsock, ByVal TextToSend As String)
     SockControl.SendData TextToSend
     DoEvents
End Sub

Sub SetConnectComputerEffect(ByRef ctlListView As ListView, ByVal IPAddress As String, ByVal OnThis As Boolean)
     Dim i          As Integer
     Dim ico        As Integer
     Dim temp()     As String
     
     If OnThis Then ico = 1 Else: ico = 2
     
     For i = 1 To ctlListView.ListItems.Count
          temp = Split(ctlListView.ListItems(i).Key, "|")
          
          If temp(1) = IPAddress Then
               ctlListView.ListItems(i).SmallIcon = ico
               
               WSAddVars(i).ListViewIndex = i
          End If
     Next i
End Sub

Function IsDaClientComputerAllowed(ByVal ctlListView As ListView, ByVal IPAddress As String) As Integer
     Dim i          As Integer
     Dim temp()     As String
     
     IsDaClientComputerAllowed = -1
     
     
     For i = 1 To ctlListView.ListItems.Count
          temp = Split(ctlListView.ListItems(i).Key, "|")
          
          If temp(1) = IPAddress Then
               IsDaClientComputerAllowed = i
               
               Exit For
          End If
     Next i
End Function
