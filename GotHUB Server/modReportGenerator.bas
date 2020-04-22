Attribute VB_Name = "modReportGenerator"
Option Explicit
' generates report via HTML

Public Enum GenerateReports
     Daily = 0
     Monthly = 1
     Yearly = 2
End Enum

Sub GenerateReport(ByVal GenRept As GenerateReports)
     Dim Header          As String
     
     Dim SQL             As String
     Dim Table           As Recordset
     
     Header = "<table width=646 border=0 align=center cellpadding=0 cellspacing=0>" + vbCrLf
     Header = Header + "<tr class=style35>" + vbCrLf
     Header = Header + "<td colspan=2>GotHUB? Internet : DAILY REPORT </td>" + vbCrLf
     Header = Header + "</tr>" + vbCrLf
     Header = Header + "<tr class='style29'>" + vbCrLf
     Header = Header + "<td width='50'>Date </td>" + vbCrLf
     Header = Header + "<td width='586'>" + Date + "</td>" + vbCrLf
     Header = Header + "</tr>" + vbCrLf
     Header = Header + "</table>" + vbCrLf
     
     If GenRept = Daily Then
          SQL = "SELECT ServicedPeoples.Name, InternetService_Details.LogInTime, InternetService_Details.LogInDate, InternetService_Details.LogOutTime, InternetService_Details.TimeUsed, Services.ServiceName, Service_Details.Quantity, Service_Details.Amount, Receipt.TotalBill "
          SQL = SQL + "FROM Services INNER JOIN (((ServicedPeoples INNER JOIN Receipt ON ServicedPeoples.SrvPipsID = Receipt.SrvPipsID) INNER JOIN InternetService_Details ON ServicedPeoples.SrvPipsID = InternetService_Details.SrvPipsID) INNER JOIN Service_Details ON ServicedPeoples.SrvPipsID = Service_Details.SrvPipsID) ON Services.ServiceID = Service_Details.ServiceID "
          SQL = SQL + "WHERE Day(Receipt.R_Date)='" + Day(Date) + "'"
     ElseIf GenRept = Monthly Then
     ElseIf GenRept = Yearly Then
     End If
End Sub
