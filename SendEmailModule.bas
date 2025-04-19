Attribute VB_Name = "SendEmailModule"
Sub Send_Email(wkNr, customer, reasonForTrip, systemType, saveDirectory, reportName, serialNumber, name, eslTotal, currencyCode)
  Dim EmailApp As Object
  Dim NewEmailItem As Object
  Dim originalBody As String
  Dim emailBody As String
  Dim firstName As String
  Dim nameBits As Variant
  Dim formattedTotal As String
  
  nameBits = Split(name, " ")
  firstName = nameBits(0)
  formattedTotal = Format(eslTotal, "0.00") & " " & currencyCode
  
  emailBody = "Hello all,<br><br>" & _
                          "Please find attached my expenses for Week " & wkNr & " at:<br><br>" & _
                          "Customer - " & customer & "<br>" & _
                          "System - " & systemType & "<br>" & _
                          "Serial - " & serialNumber & "<br>" & _
                          "Obligation - " & reasonForTrip & "<br><br>" & _
                          "<b>Total Claimed: " & formattedTotal & "</b><br><br>" & _
                          "Regards,<br><br>" & _
                          firstName
                          
  Set EmailApp = CreateObject("Outlook.Application")
  Set NewEmailItem = EmailApp.CreateItem(0)
  
  NewEmailItem.to = "ExpenseReport@icpms.com"
  NewEmailItem.CC = "creardon@icpms.com"
  NewEmailItem.Subject = "Expenses - Week " & wkNr & " - " & customer
  
  With NewEmailItem.Attachments
    .Add (saveDirectory & reportName & ".xlsx") ' Template path
    .Add (saveDirectory & reportName & ".pdf")  ' PDF path
  End With
  
  NewEmailItem.Display
  'NewEmailItem.HTMLBody = Replace(NewEmailItem.HTMLBody, "<div class=WordSection1><p class=MsoNormal><o:p>", "<div class=WordSection1><p class=MsoNormal><o:p>" & emailBody)
  NewEmailItem.HTMLBody = "<font face=""aptos"" style=""font-size:11pt;"">" & emailBody & NewEmailItem.HTMLBody & "</font>"
  
  If autoSend = "TRUE" Then
      Application.StatusBar = "Sending Email..."
      NewEmailItem.Send
  End If
  
  Set EmailApp = Nothing
  Set NewEmailItem = Nothing

End Sub

Sub Send_Summary_Email(reportsData As Collection) ' Accept the collection
  Dim EmailApp As Object
  Dim NewEmailItem As Object
  Dim emailBody As String
  Dim reportInfo As Variant ' To hold data for one report from the collection
  Dim reportDetails As String ' To build the list/table of reports
  Dim firstName As String
  Dim submitterFullName As String
  Dim nameBits As Variant
  Dim grandTotal As Double ' Variable for Grand Total
  Dim formattedIndividualTotal As String
  Dim formattedGrandTotal As String
  Dim grandTotalCurrencyCode As String ' To store the currency for the grand total

  ' Exit if there's nothing to process
  If reportsData.Count = 0 Then Exit Sub

  ' Initialize variables
  grandTotal = 0
  grandTotalCurrencyCode = "" ' Initialize currency code

  ' Get the first name from the last report processed (assuming it's the same submitter)
  If reportsData.Count > 0 Then
      submitterFullName = CStr(reportsData(reportsData.Count)(7))
      nameBits = Split(reportsData(reportsData.Count)(7), " ")
      firstName = nameBits(0)
  Else
      submitterFullName = "Unknown Submitter"
      firstName = "Team" ' Fallback if collection is somehow empty after check
  End If

  ' --- Build Initial Email Body Structure ---
  emailBody = "Hello all,<br><br>" & _
              "Please find attached my expenses for the following weeks:<br><br>"

  ' --- Start HTML Table Structure ---
  reportDetails = "<table border='1' cellpadding='5' style='border-collapse:collapse; font-family:aptos; font-size:10pt;'>" & _
                  "<thead><tr><th>Week</th><th>Customer</th><th>System</th><th>Serial</th><th>Obligation</th><th>Total</th></tr></thead>" & _
                  "<tbody>" ' Start table body

  ' --- Loop Through Each Report in the Collection ---
  For Each reportInfo In reportsData
      Dim individualTotal As Double
      Dim individualCurrencyCode As String

      ' Extract data - Assumes checkCell is index 8, repCurr is index 9
      individualTotal = reportInfo(8)
      individualCurrencyCode = reportInfo(9)

      ' Format individual total manually (e.g., "123.45 EUR")
      formattedIndividualTotal = Format(individualTotal, "0.00") & " " & individualCurrencyCode

      ' Add to grand total
      grandTotal = grandTotal + individualTotal

      ' Store the currency code of the last report processed (will be used for the grand total)
      grandTotalCurrencyCode = individualCurrencyCode

      ' Add row to the HTML table details string
      reportDetails = reportDetails & "<tr>" & _
                           "<td>" & reportInfo(0) & "</td>" & _
                           "<td>" & reportInfo(1) & "</td>" & _
                           "<td>" & reportInfo(3) & "</td>" & _
                           "<td>" & reportInfo(6) & "</td>" & _
                           "<td>" & reportInfo(2) & "</td>" & _
                           "<td style='text-align:right;'>" & formattedIndividualTotal & "</td>" & _
                           "</tr>"
  Next reportInfo

  ' --- Format the Grand Total ---
  If grandTotalCurrencyCode <> "" Then ' Check if we processed any reports
       formattedGrandTotal = Format(grandTotal, "0.00") & " " & grandTotalCurrencyCode
  Else
       formattedGrandTotal = Format(grandTotal, "0.00") ' Fallback if no currency code found
  End If

  ' --- Add the Grand Total Row to the table ---
  reportDetails = reportDetails & "<tr>" & _
                       "<td colspan='5' style='text-align:right; font-weight:bold;'>Total Claimed:</td>" & _
                       "<td style='text-align:right; font-weight:bold;'>" & formattedGrandTotal & "</td>" & _
                       "</tr>"

  ' --- Close the table body and table ---
  reportDetails = reportDetails & "</tbody></table><br>"

  ' --- Compose Final Email Body String ---
  emailBody = emailBody & reportDetails & _
              "Regards,<br><br>" & _
              firstName

  ' --- Create and Configure Email Object ---
  On Error Resume Next ' Basic error handling for Outlook automation
  Set EmailApp = CreateObject("Outlook.Application")
  If Err.Number <> 0 Then
      MsgBox "Could not create Outlook Application object. Is Outlook running/installed and configured?", vbCritical, "Outlook Error"
      Exit Sub
  End If
  Set NewEmailItem = EmailApp.CreateItem(0)
  If Err.Number <> 0 Then
       MsgBox "Could not create new Outlook email item.", vbCritical, "Outlook Error"
       Set EmailApp = Nothing
       Exit Sub
  End If
  On Error GoTo 0 ' Turn off resume next

  ' --- Configure Email Properties ---
  With NewEmailItem
      .to = "ExpenseReport@icpms.com" ' Set recipient
      .CC = "creardon@icpms.com"      ' Set CC recipient
      .Subject = submitterFullName & " - Expense Summary" ' Set subject

      ' --- Add Attachments ---
      On Error Resume Next ' Handle errors if files are missing during attachment
      For Each reportInfo In reportsData
          ' Assumes saveDirectory is index 4, reportName is index 5
          .Attachments.Add (reportInfo(4) & reportInfo(5) & ".xlsx")
          .Attachments.Add (reportInfo(4) & reportInfo(5) & ".pdf")
          If Err.Number <> 0 Then
             Debug.Print "Warning: Could not attach file for report Week " & reportInfo(0) & " - " & reportInfo(1) & ". Error: " & Err.Description
             Err.Clear ' Clear error to continue loop
          End If
      Next reportInfo
      On Error GoTo 0 ' Turn off resume next

      ' --- Set Body and Display/Send ---
      .Display ' Display the email first for review
      ' Prepend the custom body to any existing signature etc.
      .HTMLBody = "<font face=""aptos"" style=""font-size:11pt;"">" & emailBody & .HTMLBody & "</font>"

      ' Check Auto-Send Setting
      If modSettings.autoSend Then
          Application.StatusBar = "Sending Summary Email..."
          Debug.Print "Auto-sending summary email..."
          On Error Resume Next ' Handle potential send error
          .Send
          If Err.Number <> 0 Then
              MsgBox "Error occurred while trying to auto-send the summary email: " & Err.Description, vbExclamation, "Auto-Send Error"
              Err.Clear
          End If
          On Error GoTo 0
      End If
  End With

  ' --- Cleanup ---
  Set NewEmailItem = Nothing
  Set EmailApp = Nothing
  Debug.Print "Summary email created/sent."

End Sub

