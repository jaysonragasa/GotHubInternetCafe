Attribute VB_Name = "modSettings"
Option Explicit

Public Const MyKey As String = "HKLM\Software\GotHUBClient\"

Public Type Settings
     SettingsPassword              As String
     
     Server_IPAddress              As String
     server_WorkstationLtr         As String
     Server_Port                   As String
End Type

Public Setting                     As Settings

Sub InitializeSettings()
     Dim IsInstalled          As Boolean
     
     IsInstalled = CBool(ReadData("Installed"))
     
     With Setting
          .SettingsPassword = "password"
          
          .Server_IPAddress = "127.0.0.1"
          .server_WorkstationLtr = "A"
          .Server_Port = "2270"
     End With
     
     If IsInstalled Then
          With Setting
               .SettingsPassword = ReadData("SettingsPassword")
               .Server_IPAddress = ReadData("ServerIPAddress")
               .server_WorkstationLtr = ReadData("WorkstationLtr")
          End With
     Else
          With Setting
               Call WriteData("Installed", "1")
               Call WriteData("SettingsPassword", .SettingsPassword)
               Call WriteData("ServerIPAddress", .Server_IPAddress)
               Call WriteData("WorkstationLtr", .server_WorkstationLtr)
               
               .SettingsPassword = ReadData("SettingsPassword")
               .Server_IPAddress = ReadData("ServerIPAddress")
               .server_WorkstationLtr = ReadData("WorkstationLtr")
          End With
     End If
End Sub

Function ReadData(ByVal Name As String) As String
     On Error GoTo hell
     
     ReadData = WSo.regread(MyKey + Name)
     
     Exit Function
hell:
     ReadData = "0"
End Function

Function WriteData(ByVal Name As String, ByVal Data As String) As Boolean
     'On Error GoTo hell
     
     WSo.regwrite MyKey + Name, Data
     
     WriteData = True
     Exit Function
hell:
     WriteData = False
End Function
