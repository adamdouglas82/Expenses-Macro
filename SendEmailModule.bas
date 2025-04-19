Attribute VB_Name = "SendEmailModule"
Sub Send_Email(wkNr As String, customer As String, reasonForTrip As String, systemType As String, saveDirectory As String, reportName As String, serialNumber As String, name As String, autoSend As String)
  Dim EmailApp As Object
  Dim NewEmailItem As Object
  Dim originalBody As String
  Dim emailBody As String
  Dim firstName As String
  Dim nameBits As Variant
  
  nameBits = Split(name, " ")
  firstName = nameBits(0)
  
  emailBody = "Hello all,<br><br>" & _
                          "Please find attached my expenses for Week " & wkNr & " at:<br><br>" & _
                          "Customer - " & customer & "<br>" & _
                          "System - " & systemType & "<br>" & _
                          "Serial - " & serialNumber & "<br>" & _
                          "Obligation - " & reasonForTrip & "<br><br>" & _
                          "Regards,<br><br>" & _
                          firstName
                          
  Set EmailApp = CreateObject("Outlook.Application")
  Set NewEmailItem = EmailApp.CreateItem(0)
  
  NewEmailItem.To = "ExpenseReport@icpms.com"
  NewEmailItem.CC = "creardon@icpms.com"
  NewEmailItem.Subject = "Expenses - Week " & wkNr & " - " & customer
  
  With NewEmailItem.Attachments
    .Add (saveDirectory & reportName & ".xlsx") ' Template path
    .Add (saveDirectory & reportName & ".pdf")  ' PDF path
  End With
  
  NewEmailItem.Display
  'NewEmailItem.HTMLBody = Replace(NewEmailItem.HTMLBody, "<div class=WordSection1><p class=MsoNormal><o:p>", "<div class=WordSection1><p class=MsoNormal><o:p>" & emailBody)
  NewEmailItem.HTMLBody = "<font face=""aptos"" style=""font-size:11pt;"">" & emailBody & NewEmailItem.HTMLBody & "</font>"
  
  If autoSend = "Yes" Then
      Application.StatusBar = "Sending Email..."
      NewEmailItem.Send
  End If
  
  Set EmailApp = Nothing
  Set NewEmailItem = Nothing

End Sub

