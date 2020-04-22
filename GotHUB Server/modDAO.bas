Attribute VB_Name = "modDAO"
Option Explicit

Public DAO                    As Database

Public Type OverallStatus
     OA_PCUSED                As Integer
     OA_PCNOTUSED             As Integer
     OA_PCCONNECTED           As Integer
End Type

Public Type ItemDetails
     ItemIndex                As Integer
     ItemName                 As String
     ClientPCID               As String
     IPNumber                 As String
     IsConnected              As Boolean
     IsLogedIn                As Boolean
End Type

Public Type TemporaryRecord
     ClientPCID               As String
     IPNumber                 As String
     PCUser                   As String
     LogDate                  As String
     LogInTime                As String
     LogOutTime               As String
     TimeUsed                 As String
     Services                 As String      'SRV001=<QuantityNO>|SRV002=<QuantityNO>|SRV003=<QuantityNO> <-do this for multiple services
End Type

Public Type FullRecord
     Name                     As String      ' name of serviced person
     ClientPCID               As String
     IPNumber                 As String      ' ip address of the workstation used (if any)
     LogInDate                As String
     LogInTime                As String
     LogOutTime               As String
     TimeUsed                 As String
     IU_Amount                As String
     ServiceID()              As String
     QUantity()               As Integer
     Amount()                 As Currency
End Type

Public Type SelectedService
     ServiceID                As String
     Name                     As String
     Rate                     As Currency
     Unit                     As String
End Type

Public Enum GenerateReports
     Daily = 0
     Monthly = 1
     Yearly = 2
End Enum

Public OverallStat            As OverallStatus
Public ItmDetails             As ItemDetails
Public TempRec                As TemporaryRecord
Public FullRec                As FullRecord
Public SelService             As SelectedService
Public INETRent               As Currency

Public SelectedService        As String
Public SelServListIndx        As Integer

Sub InitDAO()
     On Error GoTo hell
     
     Dim s_DB_File            As String
     
     s_DB_File = FSo.buildpath(App.Path, "GotHUB_DB.mdb")
     
     Set DAO = OpenDatabase(s_DB_File, True, False, ";pwd=z3r0sl0t")
     
     Exit Sub
hell:
     MsgBox Err.Description
     
     End
End Sub

Function RecoundCount(ByVal TableName As String) As String
     Dim SQL        As String
     Dim Table      As Recordset
     
     SQL = "SELECT Count(*) FROM " & TableName
     
     Set Table = DAO.OpenRecordset(SQL)
     
     RecoundCount = Table.Fields(0).Value
     
     Set Table = Nothing: SQL = vbNullString
End Function

Sub GetWorkStationNames(ByRef ctlListView As ListView)
     Dim SQL        As String
     Dim Table      As Recordset
     Dim Items      As ListItem
     
     If Not HasRec("ClientPC") Then Exit Sub
     
     ctlListView.ListItems.Clear
     
     SQL = "SELECT * FROM ClientPC"
     Set Table = DAO.OpenRecordset(SQL)
     
     With Table
          .MoveFirst
          
          Do While Not .EOF
               Set Items = ctlListView.ListItems.Add(, Table.Fields("ClientPCID").Value + "|" + Table.Fields("IPNumber").Value, Table.Fields("PCName").Value, , 2)
               Items.Tag = 0
               .MoveNext
          Loop
     End With
     
     AutoResizeListView ctlListView
     
     Set Table = Nothing: SQL = vbNullString
End Sub

Sub GetServices(ByRef ctlListView As ListView)
     Dim SQL        As String
     Dim Table      As Recordset
     Dim Items      As ListItem
     
     If Not HasRec("Services") Then Exit Sub
     
     ctlListView.ListItems.Clear
     
     SQL = "SELECT * FROM Services"
     Set Table = DAO.OpenRecordset(SQL)
     
     With Table
          .MoveFirst
          
          Do While Not .EOF
               Set Items = ctlListView.ListItems.Add(, .Fields("ServiceID").Value, .Fields("ServiceName").Value)
               Items.SubItems(1) = .Fields("ServceAmount").Value
               Items.SubItems(2) = .Fields("PerUnit").Value
               
               If ctlListView.ColumnHeaders.Count > 3 Then
                    Items.SubItems(3) = IIf(InStr(1, UCase$(.Fields("ServiceName").Value), "INTERNET") <> 0, vbNullString, "0")
               End If
               
               .MoveNext
          Loop
     End With
     
     Set Table = Nothing: SQL = vbNullString
     
     Call AutoResizeListView(ctlListView)
End Sub

Function HasRec(ByVal TableName As String) As Boolean
     HasRec = CBool(RecoundCount(TableName))
End Function

Function DoTemporaryRecord(ByVal Add As Boolean) As Boolean
     'On Error GoTo hell
     
     Dim SQL        As String
     Dim temp       As Variant     ' splits services
     Dim temp2      As Variant     ' splits services quantity
     Dim i          As Integer
     
     With TempRec
          If Add Then
               SQL = "INSERT INTO ClientPCStatus(ClientPCID, IPNumber, PCUser, LogDate, LogInTime, LogOUtTime, TimeUsed) " & _
                     "VALUES ('" + .ClientPCID + "', " + _
                             "'" + .IPNumber + "', " + _
                             "'" + .PCUser + "', " + _
                             "'" + .LogDate + "', " + _
                             "'" + .LogInTime + "', " + _
                             "'" + .LogOutTime + "', " + _
                             "'" + .TimeUsed + "')"
                             
               DAO.Execute SQL
               DoEvents
               
               temp = Split(.Services, "|")       ' split services
               
               For i = 0 To UBound(temp)
                    temp2 = Split(temp(i), "=")   ' splits services quantity
                    
                    SQL = "INSERT INTO ServiceList(IPNumber, ServiceID, Quantity) " + _
                          "VALUES('" + .IPNumber + "', " + _
                                 "'" + temp2(0) + "', " + _
                                 "'" + temp2(1) + "')"
                                  
                    DAO.Execute SQL
                    
                    DoEvents
               Next i
          ElseIf Add = False Then    ' needs updating
               SQL = "UPDATE ClientPCStatus " + _
                     "SET " + _
                         "PCUser='" + .PCUser + "', " + _
                         "LogDate=" & Format(.LogDate, "d-m-y") & ", " + _
                         "LogInTime='" & Format(.LogInTime, "hh:mm:ss am/pm") & "', " + _
                         "LogOutTime='" & Format(.LogOutTime, "hh:mm:ss am/pm") & "', " + _
                         "TimeUsed='" + .TimeUsed + "' " + _
                     "WHERE IPNumber='" + .IPNumber + "'"

               DAO.Execute SQL
               DoEvents
          End If
     End With
     
     DoTemporaryRecord = True
     Exit Function
hell:
     DoTemporaryRecord = False
     MsgBox Err.Description, vbExclamation, "error in procedure: DoTemporaryRecord()"
End Function

Function RemoveTemporaryRecordBy(ByVal IPNumber As String) As Boolean
     Dim SQL        As String
     
     SQL = "DELETE FROM ServiceList WHERE IPNumber='" + IPNumber + "'"
     DAO.Execute SQL
     
     SQL = "DELETE FROM ClientPCStatus WHERE IPNumber='" + IPNumber + "'"
     DAO.Execute SQL
     
     SQL = vbNullString
End Function

Function GetClientServices(ByRef ctlListView As ListView, ByVal IPNumber As String)
     On Error Resume Next
     Dim SQL        As String
     Dim Table      As Recordset
     Dim i          As Integer
     
     For i = 1 To ctlListView.ListItems.Count
          ctlListView.ListItems(i).Checked = False
          
          If i > 2 Then ctlListView.ListItems(i).SubItems(3) = 0
     Next i
     
     ' check if the query has retrieved something, if it returns 0 then exit this function
     ' ---------------------------------------------------------------------------------------------------------------
     SQL = "SELECT Count(*) FROM ServiceList WHERE IPNumber='" + IPNumber + "'"
     Set Table = DAO.OpenRecordset(SQL)
     If Table.Fields(0).Value = 0 Then Exit Function
     ' --------------------------------------------------------------------------------------------------------------
     
     SQL = "SELECT * FROM ServiceList WHERE IPNumber='" + IPNumber + "'"
     Set Table = DAO.OpenRecordset(SQL)
     
     With Table
          .MoveFirst
          
          Do While Not .EOF
               For i = 1 To ctlListView.ListItems.Count
                    If ctlListView.ListItems(i).Key = .Fields("ServiceID").Value Then
                         ctlListView.ListItems(i).Checked = True
                         'ctlListView.ListItems(i).SubItems(3) = IIf(UCase$(.Fields("ServiceID").Value = "SRV001") Or UCase$(.Fields("ServiceID").Value = "SRV002"), vbNullString, .Fields("Quantity").Value)
                         ctlListView.ListItems(i).SubItems(3) = .Fields("Quantity").Value
                    End If
               Next i
               
               .MoveNext
          Loop
     End With
     
     Set Table = Nothing: SQL = vbNullString
End Function

Function AddNewService() As Boolean
     On Error GoTo hell
     Dim SQL        As String
     Dim Table      As Recordset
     
     Dim iMaxID     As Integer
     Dim sMaxID     As String
     
     SQL = "SELECT Max(ServiceID) FROM Services"
     Set Table = DAO.OpenRecordset(SQL)
     Table.Requery
     
     iMaxID = CInt(Right$(Table.Fields(0).Value, 3)) + 1
     sMaxID = "SRV" + Left$("000", 3 - Len(CStr(iMaxID))) & iMaxID
     
     Set Table = Nothing
     
     SQL = "INSERT INTO Services(ServiceID, ServiceName, ServceAmount, PerUnit) " + _
           "VALUES('" + sMaxID + "', " + _
                  "'" + SelService.Name + "', " + _
                  "'" & SelService.Rate & "', " + _
                  "'" + SelService.Unit + "')"
                  
     DAO.Execute SQL
     
     AddNewService = True
     Exit Function
hell:
     AddNewService = False
     MsgBox Err.Description, vbExclamation, "error in procedure: AddNewService()"
End Function

Function EditService() As Boolean
     On Error GoTo hell
     Dim SQL        As String
     
     With SelService
          SQL = "UPDATE Services SET ServiceName='" + .Name + "', " + _
                                    "ServceAmount='" & .Rate & "', " + _
                                    "PerUnit='" + .Unit + "' " + _
                "WHERE ServiceID='" + .ServiceID + "'"
     End With
     
     DAO.Execute SQL
     
     EditService = True
     Exit Function
hell:
     EditService = False
     MsgBox Err.Description, vbExclamation, "error in procedure: Edit Service()"
End Function

Function DeleteService() As Boolean
     On Error GoTo hell
     Dim SQL        As String
     
     SQL = "DELETE FROM Services WHERE ServiceID='" + SelService.ServiceID + "'"
     
     DAO.Execute SQL
     
     DeleteService = True
     Exit Function
hell:
     DeleteService = False
     MsgBox Err.Description, vbExclamation, "error in procedure: Edit DeleteService()"
End Function

Function AddClientServices(ByVal IPNumber As String, ByVal ctlListView As ListView)
     On Error GoTo hell
     Dim SQL        As String
     Dim Table      As Recordset
     
     Dim i          As Integer
     
     ' search
     For i = 1 To ctlListView.ListItems.Count
          ' for selected service
          If ctlListView.ListItems(i).Checked = True Then
               ' check if the selected service are already added
               SQL = "SELECT Count(*) " + _
                     "FROM ServiceList " + _
                     "WHERE ServiceList.IPNumber='" + IPNumber + "' AND " + _
                           "ServiceList.ServiceID='" + ctlListView.ListItems(i).Key + "'"
                           
               Set Table = DAO.OpenRecordset(SQL)
               
               ' if not then add the selected service
               If Table.Fields(0).Value = 0 Then
                    SQL = "INSERT INTO ServiceList(IPNumber, ServiceID, Quantity) " + _
                          "VALUES ('" + IPNumber + "', " + _
                                  "'" + ctlListView.ListItems(i).Key + "', " + _
                                  "" & CSng(ctlListView.ListItems(i).SubItems(3)) & ")"
                    
                    DAO.Execute SQL
                    DoEvents
               
               ' if found then just update the quantity even if not requested
               ElseIf Table.Fields(0).Value <> 0 Then
                    If ctlListView.ListItems(i).SubItems(3) <> vbNullString Then
                         SQL = "UPDATE ServiceList " + _
                               "SET Quantity=" & CSng(ctlListView.ListItems(i).SubItems(3)) & " " + _
                               "WHERE ServiceList.IPNumber='" + IPNumber + "' AND " + _
                                     "ServiceList.ServiceID='" + ctlListView.ListItems(i).Key + "'"
                                
                         DAO.Execute SQL
                    End If
               End If
          End If
     Next i
     
     Set Table = Nothing
     SQL = vbNullString
     
     AddClientServices = True
     Exit Function
hell:
     AddClientServices = False
     MsgBox Err.Description, vbExclamation, "error in procedure: AddClientServices()"
End Function

Function RemoveSelectedService(ByVal IPNumber As String, ByVal ServiceID As String) As Boolean
     On Error GoTo hell
     Dim SQL        As String
     
     SQL = "DELETE FROM ServiceList WHERE ServiceList.IPNumber='" + IPNumber + "' and ServiceList.ServiceID='" + ServiceID + "'"
     DAO.Execute SQL
     DoEvents
     
     RemoveSelectedService = True
     Exit Function
hell:
     RemoveSelectedService = False
     MsgBox Err.Description, vbExclamation, "error in procedure: RemoveSelectedService()"
End Function

Function ChangeInternetService(ByVal IPNumber As String, ByVal ServiceID As String) As Boolean
     On Error GoTo hell
     Dim SQL        As String
     
     SQL = "UPDATE ServiceList SET ServiceID='" + ServiceID + "' WHERE ServiceList.IPNumber='" + IPNumber + "' AND (ServiceList.ServiceID='SRV001' OR ServiceList.ServiceID='SRV002')"
     DAO.Execute SQL
     
     ChangeInternetService = True
     Exit Function
hell:
     ChangeInternetService = False
     MsgBox Err.Description, vbExclamation, "error in procedure: ChangeInternetService()"
End Function

Function GetClientStatus(ByRef ctlListView As ListView, ByVal SelIndex As Integer, ByVal IPNumber As String) As String
     On Error Resume Next
     Dim SQL        As String
     Dim Table      As Recordset
     Dim temp()     As String
     
     Dim i          As Integer
     Dim c          As Integer
     
     ' check if the query has retrieved something, if it returns 0 then exit this function
     ' ---------------------------------------------------------------------------------------------------------------
     'SQL = "SELECT Count(*) FROM ClientPCStatus WHERE ClientPCStatus.IPNumber='" + IPNumber + "'"
     'Set Table = DAO.OpenRecordset(SQL)
     'If Table.Fields(0).Value = 0 Then Exit Function
     ' --------------------------------------------------------------------------------------------------------------
     
     SQL = "SELECT IPNumber, PCUser, LogInTime, LogOutTime, TimeUsed FROM ClientPCStatus WHERE IPNumber='" + IPNumber + "'"
     Set Table = DAO.OpenRecordset(SQL)
     Table.Requery
     
     If Table.RecordCount = 0 Then
          GetClientStatus = 0
     
          Exit Function
     End If
     
     With Table
          .MoveFirst
          
          'Do While Not .EOF
               For i = 1 To ctlListView.ListItems.Count
                    temp = Split(ctlListView.ListItems(i).Key, "|")
                    
                    If temp(1) = .Fields("IPNumber").Value Then
                         For c = 1 To .Fields.Count
                              ctlListView.ListItems(i).SubItems(c) = .Fields(c).Value
                         Next c
                         
                         ctlListView.ListItems(i).Tag = 1
                         
                         'If frmMain.sckServer(WSAddVars(i).WinsockIndex).State = sckConnected Then
                              'frmMain.sckServer(WSAddVars(i).WinsockIndex).SendData CONN_LOGIN: DoEvents
                              'frmMain.sckServer(WSAddVars(i).WinsockIndex).SendData "TIMER|" & Time & "|" & Left$(ctlListView.ListItems(i).SubItems(3), 7): DoEvents
                              'MsgBox "wen iso"
                              'MsgBox "UNAME|" + ctlListView.ListItems(i).SubItems(1)
                              'frmMain.sckServer(WSAddVars(i).WinsockIndex).SendData "UNAME|" & ctlListView.ListItems(i).SubItems(1): DoEvents
                         'End If
                         
                         GetClientStatus = i
                         
                         Exit For
                    End If
               Next i
               
               '.MoveNext
          'Loop
     End With
     
     Set Table = Nothing: SQL = vbNullString
End Function

Function ChangeWorkstation(ByVal OldIPNumber As String, ByVal NewIPNumber As String) As Boolean
     On Error GoTo hell
     Dim SQL        As String
     
     SQL = "UPDATE ClientPCStatus, ServiceList " + _
           "SET ClientPCStatus.IPNumber='" + NewIPNumber + "', " + _
               "ServiceList.IPNumber='" + NewIPNumber + "' " + _
           "WHERE ClientPCStatus.IPNumber='" + OldIPNumber + "' AND " + _
                 "ServiceList.IPNumber='" + OldIPNumber + "'"
                 
     DAO.Execute SQL
     
     ChangeWorkstation = True
     Exit Function
hell:
     ChangeWorkstation = False
     MsgBox Err.Description, vbExclamation, "error in procedure: ChangeWorkstation()"
End Function

'Public Type FullRecord
'     Name                     As String      ' name of serviced person
'     IPNumber                 As String      ' ip address of the workstation used (if any)
'     LogInDate                As String
'     LogInTime                As String
'     LogOutTime               As String
'     TimeUsed                 As String
'--------------------------------------------------------------------------------------------------------------------
'     ServiceID()              As String
'     QUantity()               As Integer
'     Amount()                 As Currency
'End Type

'SPIU00000      unique id for ServicedPeoplesWhoUsedTheInternet
'SPD00000       unique id for ServicedPeopleDetails
'SP00000        unique id for ServicedPeoples

Function MaxID(ByVal TableName As String, ByVal PreFix As String, ByVal UqID As String) As String
     Dim SQL        As String
     Dim Table      As Recordset
     
     Dim ret        As String
     Dim temp       As String
     Dim MxID       As Integer
     
     Dim i
     Dim Zero       As String
     
     Zero = "00000"
     
     If PreFix = "ORNO" Then Zero = "0000000000"
     
     SQL = "SELECT Count(*) FROM " + TableName
     Set Table = DAO.OpenRecordset(SQL)
     
     If Table.Fields(0).Value = 0 Then
          ret = Left$(Zero, Len(Zero) - 1) + "1"
     Else
          SQL = "SELECT Max(" + UqID + ") FROM " + TableName
          Set Table = DAO.OpenRecordset(SQL)
     
          temp = Right$(Table.Fields(0).Value, Len(Table.Fields(0).Value) - Len(PreFix))
          MxID = CInt(temp) + 1
          ret = Left$(temp, Len(temp) - Len(CStr(MxID))) + CStr(MxID)
     End If
     
     MaxID = PreFix + ret
     
     Set Table = Nothing
End Function

Function CreateRecord() As Boolean
     On Error GoTo hell
     Dim SQL             As String
     Dim i               As Integer
     Dim MaxSrvPipsID    As String
     Dim TtlBil          As Currency
     
     MaxSrvPipsID = MaxID("ServicedPeoples", "SP", "SrvPipsID")
     
     SQL = "INSERT INTO ServicedPeoples(SrvPipsID, Name) " + _
           "VALUES('" + MaxSrvPipsID + "', " + _
                  "'" + FullRec.Name + "')"
     DAO.Execute SQL
     DoEvents
     
     SQL = "INSERT INTO InternetService_Details(SrvPips_IU_ID, SrvPipsID, ClientPCID, LogInDate, LogInTime, LogOutTime, TimeUsed, TotalAmount) " + _
           "VALUES('" + MaxID("InternetService_Details", "SPIU", "SrvPips_IU_ID") + "', " + _
                  "'" + MaxSrvPipsID + "', " + _
                  "'" + FullRec.ClientPCID + "', " + _
                  "'" & Date & "', " + _
                  "'" + FullRec.LogInTime + "', " + _
                  "'" + FullRec.LogOutTime + "', " + _
                  "'" + FullRec.TimeUsed + "', " + _
                  "'" + FullRec.IU_Amount + "')"
                  
     DAO.Execute SQL
     DoEvents
     
     TtlBil = FullRec.IU_Amount
     
     For i = 0 To UBound(FullRec.ServiceID)
          SQL = "INSERT INTO Service_Details(SrvPipsDetID, SrvPipsID, ServiceID, Quantity, Amount) " + _
                "VALUES('" + MaxID("Service_Details", "SPD", "SrvPipsDetID") + "', " + _
                       "'" + MaxSrvPipsID + "', " + _
                       "'" & FullRec.ServiceID(i) & "', " + _
                       "'" & FullRec.QUantity(i) & "', " + _
                       "'" & FullRec.Amount(i) & "')"
          
          TtlBil = TtlBil + FullRec.Amount(i)
          
          DAO.Execute SQL
          DoEvents
     Next i
     
     SQL = "INSERT INTO Receipt(ORNO, SrvPipsID, R_Date, R_Time, TotalBill) " + _
           "VALUES('" + MaxID("Receipt", vbNullString, "ORNO") + "', " + _
                  "'" + MaxSrvPipsID + "', " + _
                  "'" & Date & "', " + _
                  "'" + FullRec.LogOutTime + "', " + _
                  "'" & TtlBil & "')"
     DAO.Execute SQL
     DoEvents
     
     CreateRecord = True
     Exit Function
hell:
     CreateRecord = False
     MsgBox Err.Description
End Function

Function CreateWalkthroughRecord() As Boolean
     On Error GoTo hell
     
     Dim i               As Integer
     Dim SQL             As String
     Dim MaxSrvPipsID    As String
     Dim TtlBil          As Currency
     Dim sMaxID          As String
     Dim sWSMaxID        As String
     
     MaxSrvPipsID = MaxID("ServicedPeoples", "SP", "SrvPipsID")
     
     SQL = "INSERT INTO ServicedPeoples(SrvPipsID, Name) " + _
           "VALUES('" + MaxSrvPipsID + "', " + _
                  "'" + FullRec.Name + "')"
                  
     DAO.Execute SQL: DoEvents
     
     For i = 0 To UBound(FullRec.ServiceID)
          sMaxID = MaxID("Service_Details", "SPD", "SrvPipsDetID")
          sWSMaxID = MaxID("WalkthroughService_Details", "WP", "WSID")
          
          SQL = "INSERT INTO Service_Details(SrvPipsDetID, SrvPipsID, ServiceID, Quantity, Amount) " + _
                "VALUES('" + sMaxID + "', " + _
                       "'" + MaxSrvPipsID + "', " + _
                       "'" & FullRec.ServiceID(i) & "', " + _
                       "'" & FullRec.QUantity(i) & "', " + _
                       "'" & FullRec.Amount(i) & "')"
          
          TtlBil = TtlBil + FullRec.Amount(i)
          
          DAO.Execute SQL
          DoEvents
          
          SQL = "INSERT INTO WalkthroughService_Details(WSID, SrvPipsDetID) " + _
                "VALUES('" + sWSMaxID + "', '" + sMaxID + "')"
                
          DAO.Execute SQL
          DoEvents
     Next i
     
     SQL = "INSERT INTO Receipt(ORNO, SrvPipsID, R_Date, R_Time, TotalBill) " + _
           "VALUES('" + MaxID("Receipt", vbNullString, "ORNO") + "', " + _
                  "'" + MaxSrvPipsID + "', " + _
                  "'" & Date & "', " + _
                  "'" + FullRec.LogOutTime + "', " + _
                  "'" & TtlBil & "')"
     DAO.Execute SQL
     DoEvents
     
     CreateWalkthroughRecord = True
     Exit Function
hell:
     CreateWalkthroughRecord = False
     MsgBox Err.Description
End Function

Function Update_WS_IPNumber(ByVal IPNumber As String, ByVal NewWorkstationName As String, ByVal NewIPNumber As String) As Boolean
     On Error GoTo hell
     Dim SQL        As String
     
     SQL = "UPDATE ClientPC SET ClientPC.IPNumber='" + NewIPNumber + "', ClientPC.PCName='" + NewWorkstationName + "' WHERE ClientPC.IPNumber='" + IPNumber + "'"
     DAO.Execute SQL
     
     Update_WS_IPNumber = True
     Exit Function
hell:
     Update_WS_IPNumber = False
     MsgBox Err.Description, vbExclamation, "error in procedure: UpdateComputerSettings()"
End Function

Sub GenerateReport(ByVal GenRept As Integer, Optional DateNow As String)
     Dim Header          As String
     Dim Datas           As String
     
     Dim SQLNames        As String
     Dim SQLServis       As String
     Dim Table           As Recordset
     Dim TableServ       As Recordset
     
     Dim vFile           As Variant
     
     Dim Total           As Currency
     Dim i               As Integer
     Dim Max             As Integer
     
     Dim title           As String
     
     If GenRept = 1 Then title = "Daily Report"
     If GenRept = 2 Then title = "Monthly Report"
     If GenRept = 3 Then title = "Yearly Report"
     
     Header = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN'" + vbCrLf
     Header = Header + "'http://www.w3.org/TR/html4/loose.dtd'>" + vbCrLf
     Header = Header + "<html>" + vbCrLf
     Header = Header + "<head>" + vbCrLf
     Header = Header + "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" + vbCrLf
     Header = Header + "<title>GotHUB? Internet - " + title + "</title>" + vbCrLf
     Header = Header + "<style type='text/css'>" + vbCrLf
     Header = Header + "<!--" + vbCrLf
     Header = Header + ".style27 {font-family: Arial, Helvetica, sans-serif; color: #A1A1A1; font-weight: bold; font-size: small; }" + vbCrLf
     Header = Header + ".style29 {font-family: Arial, Helvetica, sans-serif; font-size: small; }" + vbCrLf
     Header = Header + ".style35 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; }" + vbCrLf
     Header = Header + "-->" + vbCrLf
     Header = Header + "</style>" + vbCrLf
     Header = Header + "</head>" + vbCrLf
     Header = Header + "<body>" + vbCrLf
     
     Header = Header + "<table width='646' border='0' align='center' cellpadding='0' cellspacing='0'>" + vbCrLf
     Header = Header + "<tr class='style35'>" + vbCrLf
     Header = Header + "<td colspan='2'>GotHUB? Internet - " + title + "</td>" + vbCrLf
     Header = Header + "</tr>" + vbCrLf
     Header = Header + "<tr class='style29'>" + vbCrLf
     Header = Header + "<td width='50'>Date </td>" + vbCrLf
     Header = Header + "<td width='586'>" & DateNow & "</td>" + vbCrLf
     Header = Header + "</tr>" + vbCrLf
     Header = Header + "</table><br>" + vbCrLf
     
     If GenRept = 1 Then
          'SQLNames = "IU_GetNames_Now"
          SQLNames = "SELECT ServicedPeoples.SrvPipsID, ServicedPeoples.Name, InternetService_Details.LogInTime, InternetService_Details.LogInDate, InternetService_Details.LogOutTime, InternetService_Details.TimeUsed, Receipt.TotalBill "
          SQLNames = SQLNames + "FROM (ServicedPeoples INNER JOIN Receipt ON ServicedPeoples.SrvPipsID = Receipt.SrvPipsID) INNER JOIN InternetService_Details ON ServicedPeoples.SrvPipsID = InternetService_Details.SrvPipsID "
          SQLNames = SQLNames + "WHERE ((Day((InternetService_Details.LogInDate)) = " & Day(Format(DateNow, "dd-mmm-yy")) & ")) And " + _
                                      "((Month((InternetService_Details.LogInDate)) = " & Month(Format(DateNow, "dd-mmm-yy")) & "))"
          
          Set Table = DAO.OpenRecordset(SQLNames)
          Table.Requery
          DoEvents
          
          Datas = "<p class='style35' align='center'>Internet Rental</p><br>"
          Datas = Datas + "<table width='646' border='0' align='center' cellpadding='0' cellspacing='0' bordercolor='#CCCCCC'>" + vbCrLf
          
          With Table
               If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    Do While Not .EOF
                         SQLServis = "SELECT Service_Details.SrvPipsID, Services.ServiceID, Services.ServiceName, Service_Details.Quantity, Service_Details.Amount, InternetService_Details.TotalAmount "
                         SQLServis = SQLServis + "FROM (ServicedPeoples INNER JOIN InternetService_Details ON ServicedPeoples.SrvPipsID = InternetService_Details.SrvPipsID) INNER JOIN (Services INNER JOIN Service_Details ON Services.ServiceID = Service_Details.ServiceID) ON ServicedPeoples.SrvPipsID = Service_Details.SrvPipsID "
                         SQLServis = SQLServis + "WHERE (((Service_Details.SrvPipsID)='" + .Fields("SrvPipsID") + "'))"

                         Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                         Datas = Datas + "<td>Name </td>" + vbCrLf
                         Datas = Datas + "<td>LogIn Time </td>" + vbCrLf
                         Datas = Datas + "<td>LogIn Date </td>" + vbCrLf
                         Datas = Datas + "<td>LogOut Time </td>" + vbCrLf
                         Datas = Datas + "<td>Time Used </td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         
                         Datas = Datas + "<tr class='style29'>" + vbCrLf
                         Datas = Datas + "<td>" + .Fields("Name") + " </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("LogInTime") & " </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("LogInDate") & " </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("LogOutTime") & " </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("TimeUsed") & " </td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         
                         Datas = Datas + "<tr class='style27'>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Services </td>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Service Name </td>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Quantity </td>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Total Amount </td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                                                
                         Set TableServ = DAO.OpenRecordset(SQLServis)
                         TableServ.Requery
                         DoEvents
                         
                         If TableServ.RecordCount <> 0 Then
                              TableServ.MoveFirst
                              
                              Do While Not TableServ.EOF
                                   Datas = Datas + "<tr class='style29'>" + vbCrLf
                                   Datas = Datas + "<td bgcolor='#FFFFFF'>&nbsp;</td>" + vbCrLf
                                   Datas = Datas + "<td bgcolor='#FFFFFF'>&nbsp;</td>" + vbCrLf
                                   Datas = Datas + "<td>" + TableServ.Fields("ServiceName") + " </td>" + vbCrLf
                                   Datas = Datas + "<td>" & TableServ.Fields("Quantity") & " </td>" + vbCrLf
                                   If TableServ.Fields("ServiceID") = "SRV001" Or TableServ.Fields("ServiceID") = "SRV002" Then
                                        Datas = Datas + "<td>" & TableServ.Fields("TotalAmount") & " </td>" + vbCrLf
                                        Total = Total + CCur(TableServ.Fields("TotalAmount"))
                                   Else
                                        Datas = Datas + "<td>" & TableServ.Fields("Amount") & " </td>" + vbCrLf
                                        Total = Total + CCur(TableServ.Fields("Amount"))
                                   End If
                                   
                                   Datas = Datas + "</tr>" + vbCrLf
                                   
                                   TableServ.MoveNext
                              Loop
                         End If
                         
                         Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                         Datas = Datas + "<td>Total </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("TotalBill") & " </td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         Datas = Datas + "<tr bgcolor='#FFFFFF' class='style27'>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf

                         .MoveNext
                    Loop
               End If
          End With
          
          Datas = Datas + "<tr bgcolor='#FFFFFF' class='style27'>" + vbCrLf
          Datas = Datas + "<td>Total </td>" + vbCrLf
          Datas = Datas + "<td>" & Total & "</td>" + vbCrLf
          Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
          Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
          Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
          Datas = Datas + "</tr>" + vbCrLf
          
          Datas = Datas + "</table><br>"
          Datas = Datas + "<p class='style35' align='center'>Walkthrough Service</p><br>"
          Datas = Datas + "<table width='646' border='0' align='center' cellpadding='0' cellspacing='0' bordercolor='#CCCCCC'>" + vbCrLf
          
          Total = 0
          
          'SQLNames = "WalkthroughService_Now"
          SQLNames = "SELECT ServicedPeoples.Name, Receipt.R_Date, Receipt.R_Time, Services.ServiceName, Service_Details.Quantity, Service_Details.Amount, Receipt.TotalBill, (SELECT Count([WSID]) AS Expr1 FROM Service_Details INNER JOIN WalkthroughService_Details ON Service_Details.SrvPipsDetID = WalkthroughService_Details.SrvPipsDetID WHERE (((Service_Details.SrvPipsID)=(ServicedPeoples.SrvPipsID)));) AS iMax "
          SQLNames = SQLNames + "FROM Services INNER JOIN ((ServicedPeoples INNER JOIN Receipt ON ServicedPeoples.SrvPipsID = Receipt.SrvPipsID) INNER JOIN (Service_Details INNER JOIN WalkthroughService_Details ON Service_Details.SrvPipsDetID = WalkthroughService_Details.SrvPipsDetID) ON ServicedPeoples.SrvPipsID = Service_Details.SrvPipsID) ON Services.ServiceID = Service_Details.ServiceID "
          SQLNames = SQLNames + "WHERE ((Day((Receipt.R_Date))=" & Day(Format(DateNow, "dd-mmm-yy")) & ")) And " + _
                                      "((Month((Receipt.R_Date))=" & Month(Format(DateNow, "dd-mmm-yy")) & "))"
          Set Table = DAO.OpenRecordset(SQLNames)
          Table.Requery
          DoEvents
          
          With Table
               If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    Do While Not .EOF
                         Max = .Fields("iMax").Value
                         
                         Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                         Datas = Datas + "<td>Name </td>" + vbCrLf
                         Datas = Datas + "<td>Time </td>" + vbCrLf
                         Datas = Datas + "<td>Date </td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         
                         Datas = Datas + "<tr class='style29'>" + vbCrLf
                         Datas = Datas + "<td>" + .Fields("Name") + " </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("R_Date") & " </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("R_Time") & " </td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         
                         Datas = Datas + "<tr class='style27'>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Services </td>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Service Name </td>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Quantity </td>" + vbCrLf
                         Datas = Datas + "<td bgcolor='#E1E1E1'>Total Amount </td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         
                         For i = 1 To Max
                              Datas = Datas + "<tr class='style29'>" + vbCrLf
                              Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                              Datas = Datas + "<td>" + .Fields("ServiceName") + " </td>" + vbCrLf
                              Datas = Datas + "<td>" & .Fields("Quantity") & " </td>" + vbCrLf
                              Datas = Datas + "<td>" & .Fields("Amount") & " </td>" + vbCrLf
                              Datas = Datas + "</tr>" + vbCrLf
                              
                              Total = Total + CCur(.Fields("Amount"))
                              
                              .MoveNext
                         Next i
                         
                         .MovePrevious
                         
                         Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                         Datas = Datas + "<td>Total </td>" + vbCrLf
                         Datas = Datas + "<td>" & .Fields("TotalBill") & " </td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         Datas = Datas + "<tr bgcolor='#FFFFFF' class='style27'>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                         Datas = Datas + "</tr>" + vbCrLf
                         
                         .MoveNext
                    Loop
               End If
          End With
          
          Datas = Datas + "<tr bgcolor='#FFFFFF' class='style27'>" + vbCrLf
          Datas = Datas + "<td>Total </td>" + vbCrLf
          Datas = Datas + "<td>" & Total & "</td>" + vbCrLf
          Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
          Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
          Datas = Datas + "</tr>" + vbCrLf
          
          Datas = Datas + "</table></body></html>"
     ElseIf GenRept = 2 Then
          SQLNames = "SELECT Receipt.R_Date, Sum(Receipt.TotalBill) AS SumOfTotalBill "
          SQLNames = SQLNames + "From Receipt "
          SQLNames = SQLNames + "GROUP BY Receipt.R_Date"
          
          Datas = "<table width='646' border='0' align='center' cellpadding='0' cellspacing='0' bordercolor='#CCCCCC'>" + vbCrLf
          
          Set Table = DAO.OpenRecordset(SQLNames)
          Table.Requery
          
          With Table
               If .RecordCount <> 0 Then
                    Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                    Datas = Datas + "<td>Month </td>" + vbCrLf
                    Datas = Datas + "<td>Day </td>" + vbCrLf
                    Datas = Datas + "<td>Total </td>" + vbCrLf
                    Datas = Datas + "</tr>" + vbCrLf
                    
                    Do While Not .EOF
                         If Month(DateNow) = Month(.Fields("R_Date")) Then
                              Datas = Datas + "<tr class='style29'>" + vbCrLf
                              Datas = Datas + "<td>" & MonthName(Month(DateNow)) & " </td>" + vbCrLf
                              Datas = Datas + "<td>" & Day(.Fields("R_Date")) & " </td>" + vbCrLf
                              Datas = Datas + "<td>Php " & FormatNumber(.Fields("SumOfTotalBill"), 2) & " </td>" + vbCrLf
                              Datas = Datas + "</tr>" + vbCrLf
                         
                              Total = Total + .Fields("SumOfTotalBill")
                         End If
                         
                         .MoveNext
                    Loop
                    
                    Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                    Datas = Datas + "<td>Total </td>" + vbCrLf
                    Datas = Datas + "<td>" & Total & " </td>" + vbCrLf
                    Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                    Datas = Datas + "</tr>" + vbCrLf
                    
                    .MoveFirst
               End If
               
               Datas = Datas + "</table></body></html>"
          End With
     ElseIf GenRept = 3 Then
          SQLNames = "SELECT Month([R_Date]) AS Expr1, Sum(Receipt.TotalBill) AS SumOfTotalBill "
          SQLNames = SQLNames + "From Receipt "
          SQLNames = SQLNames + "GROUP BY Month([R_Date])"
          
          Datas = "<table width='646' border='0' align='center' cellpadding='0' cellspacing='0' bordercolor='#CCCCCC'>" + vbCrLf
          
          Set Table = DAO.OpenRecordset(SQLNames)
          Table.Requery
          
          With Table
               If .RecordCount <> 0 Then
                    Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                    Datas = Datas + "<td>Year </td>" + vbCrLf
                    Datas = Datas + "<td>Month </td>" + vbCrLf
                    Datas = Datas + "<td>Total </td>" + vbCrLf
                    Datas = Datas + "</tr>" + vbCrLf
                    
                    Do While Not .EOF
                         'If Year(DateNow) = Year(.Fields("R_Date")) Then
                              Datas = Datas + "<tr class='style29'>" + vbCrLf
                              Datas = Datas + "<td>" & Year(DateNow) & " </td>" + vbCrLf
                              Datas = Datas + "<td>" & MonthName(.Fields("Expr1")) & " </td>" + vbCrLf
                              Datas = Datas + "<td>Php " & FormatNumber(.Fields("SumOfTotalBill"), 2) & " </td>" + vbCrLf
                              Datas = Datas + "</tr>" + vbCrLf
                         
                              Total = Total + .Fields("SumOfTotalBill")
                         'End If
                         
                         .MoveNext
                    Loop
                    
                    Datas = Datas + "<tr bgcolor='#E1E1E1' class='style27'>" + vbCrLf
                    Datas = Datas + "<td>Total </td>" + vbCrLf
                    Datas = Datas + "<td>" & Total & " </td>" + vbCrLf
                    Datas = Datas + "<td>&nbsp;</td>" + vbCrLf
                    Datas = Datas + "</tr>" + vbCrLf
                    
                    .MoveFirst
               End If
               
               Datas = Datas + "</table></body></html>"
          End With
     End If
     
     Set vFile = FSo.CreateTextFile(FSo.buildpath(dTmp, "tmpGH_Rpt.html"))
     vFile.write Header + Datas
     vFile.Close
     DoEvents
     
     WSo.run FSo.buildpath(dTmp, "tmpGH_Rpt.html"), vbMaximizedFocus
     
     Set vFile = Nothing
     Set Table = Nothing
     Set TableServ = Nothing
End Sub

Function GetServiceAmount(ByVal ServiceID As String) As Currency
     Dim SQL        As String
     Dim Table      As Recordset
     
     SQL = "SELECT ServceAmount FROM Services WHERE ServiceID='" + ServiceID + "'"
     Set Table = DAO.OpenRecordset(SQL)
     Table.Requery
     
     GetServiceAmount = CCur(Table.Fields(0).Value)
End Function

Function GetServiceAmount_IP(ByVal IPAddress As String) As Currency
     Dim SQL        As String
     Dim Table      As Recordset
     
     SQL = "SELECT ClientPCStatus.IPNumber, ServiceList.ServiceID "
     SQL = SQL + "FROM ClientPCStatus INNER JOIN ServiceList ON ClientPCStatus.IPNumber = ServiceList.IPNumber "
     SQL = SQL + "WHERE (((ClientPCStatus.IPNumber)='" + IPAddress + "'))"
     Set Table = DAO.OpenRecordset(SQL)
     Table.Requery
     
     GetServiceAmount_IP = GetServiceAmount(Table.Fields(1).Value)
End Function
