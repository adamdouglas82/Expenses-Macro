Attribute VB_Name = "ExpensifyAPIModule"
Function expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)

Dim postdata As String
Dim httpRequest As Object
Dim templateContent As String
Dim templateContentURL As String
Dim fullSend As String
Dim expensifyRequestReturn As String
Dim fileType As String
Dim reportFilters As String
Dim reportState As String
Dim previousYear As Integer
Dim previousYearString As String

If requestType = "csv" Then
    ' Set the URL and data to download each csv
    fileType = "csv"
      postdata = "requestJobDescription={" _
                                        & """type"":""file""," _
                                        & """credentials"":{" _
                                        & """partnerUserID"":""" & userID & """," _
                                       & """partnerUserSecret"":""" & userSecret & """}," _
                                       & """onReceive"":{""immediateResponse"":[""returnRandomFileName""]}," _
                                       & """inputSettings"":{" _
                                            & """type"":""combinedReportData""," _
                                           & """filters"":{" _
                                               & """reportIDList"": """ & reportID & """}}," _
                                            & """outputSettings"":{" _
                                               & """fileExtension"":""" & fileType & """}}"

        templateContent = "<#if addHeader == true>Submitted by:;Date;Description;Meal Members/Notes;Category;Expense Amount;Expense Currency;Exchange Rate;Converted Amount;Report Currency;Mileage;Mileage Rate;Mileage Unit;Customer;Week;Reason for Trip;System Type;Company;Serial Number;Report Total;Fees;Policy;Policy ID;Report ID<#lt></#if><#list reports as report><#list report.transactionList as expense>" & vbCrLf & _
                          "<#if expense.modifiedMerchant?has_content><#assign merchant = expense.modifiedMerchant><#else><#assign merchant = expense.merchant></#if><#if expense.convertedAmount?has_content><#assign convertedAmount = expense.convertedAmount/100></#if><#if expense.modifiedAmount?has_content><#assign amount = expense.modifiedAmount/100><#else><#assign amount = expense.amount/100></#if><#if expense.modifiedCreated?has_content><#assign created = expense.modifiedCreated><#else><#assign created = expense.created></#if>" & vbCrLf & _
                          "<#assign reportTotal = report.total/100><#assign mileageRate = expense.units.rate>" & vbCrLf & "${report.submitter.fullName};<#t>${created};<#t>${merchant};<#t>${expense.comment};<#t>${expense.category};<#t>${amount};<#t>${expense.currency};<#t>${expense.currencyConversionRate};<#t>${convertedAmount};<#t>${report.currency};<#t>${expense.units.count};<#t>${mileageRate};<#t>${expense.units.unit};<#t>${report.customField.Customer};<#t>${report.customField.Week};<#t>${report.customField.Reason_for_Trip};<#t>${report.customField.System_Type};<#t>${report.customField.Company};<#t>${report.customField.Serial_Number};<#t>${reportTotal};<#t>${report.customField.Fees};<#t>${report.policyName};<#t>${report.policyID};<#t>${report.reportID}<#lt></#list></#list>"
    
ElseIf requestType = "combinedReports" Then
    ' Set the URL and data to download which reports are closed (add check to see if they've already been downloaded "markedAsExported" with string to identify)

    fileType = "csv"
    
    ' download archived (closed) and reimbursed expenses from (Year(Now) - 1)-01-01
    ' available options are (in brackets website status): "OPEN", "SUBMITTED" (Processing), "APPROVED", "REIMBURSED", "ARCHIVED (Closed)"
    ' set to show that expenses have been exported on the expensify website
    
    previousYear = Year(Now) - 1
    previousYearString = CStr(previousYear) & "-01-01"
    
    reportFilters = """startDate"":""" & previousYearString & """}},"
    reportState = """reportState"":""ARCHIVED,REIMBURSED"","
    
    postdata = "requestJobDescription={" _
                                        & """type"":""file""," _
                                        & """credentials"":{" _
                                        & """partnerUserID"":""" & userID & """," _
                                       & """partnerUserSecret"":""" & userSecret & """}," _
                                       & """onReceive"":{""immediateResponse"":[""returnRandomFileName""]}," _
                                       & """inputSettings"":{" _
                                            & """type"":""combinedReportData""," _
                                           & reportState _
                                           & """filters"":{" _
                                               & reportFilters _
                                            & """outputSettings"":{" _
                                               & """fileExtension"":""" & fileType & """," _
                                               & """includeFullPageReceiptsPdf"":""true""}," _
                                               & """onFinish"":[{""actionName"":""markAsExported"",""label"":""Expensify Export""}]}"

    templateContent = "<#compress><#list reports as report><#assign reportTotal = report.total/100>${report.reportID};${report.status};${report.customField.Customer};${report.customField.Week};${report.submitted};${report.created};${reportTotal}" & vbCrLf & "</#list></#compress>"

ElseIf requestType = "checkReimbursed" Then
    Application.StatusBar = "Requesting list of reimbursed Reports from Expensify..."
    ' Set the URL and data to download which reports are closed (add check to see if they've already been downloaded "markedAsExported" with string to identify)
    Debug.Print "Requesting list of reimbursed Reports from Expensify..."
    fileType = "csv"
    
    reportFilters = """startDate"":""2024-01-10""}},"

    postdata = "requestJobDescription={" _
                                        & """type"":""file""," _
                                        & """credentials"":{" _
                                        & """partnerUserID"":""" & userID & """," _
                                       & """partnerUserSecret"":""" & userSecret & """}," _
                                       & """onReceive"":{""immediateResponse"":[""returnRandomFileName""]}," _
                                       & """inputSettings"":{" _
                                            & """type"":""combinedReportData""," _
                                           & """reportState"":""REIMBURSED""," _
                                           & """filters"":{" _
                                               & reportFilters _
                                            & """outputSettings"":{" _
                                               & """fileExtension"":""" & fileType & """," _
                                               & """includeFullPageReceiptsPdf"":""true""}}"
  


    templateContent = "<#compress><#list reports as report>${report.reportID};${report.status};${report.customField.Customer};${report.customField.Week};${report.submitted}" & vbCrLf & "</#list></#compress>"

ElseIf requestType = "PDF" Then

    ' Set the URL and data to download each pdf
    fileType = "pdf"
    postdata = "requestJobDescription={" _
                                        & """type"":""file""," _
                                        & """credentials"":{" _
                                        & """partnerUserID"":""" & userID & """," _
                                       & """partnerUserSecret"":""" & userSecret & """}," _
                                       & """onReceive"":{""immediateResponse"":[""returnRandomFileName""]}," _
                                       & """inputSettings"":{" _
                                            & """type"":""combinedReportData""," _
                                           & """filters"":{" _
                                               & """reportIDList"": """ & reportID & """}}," _
                                            & """outputSettings"":{" _
                                               & """fileExtension"":""" & fileType & """," _
                                               & """includeFullPageReceiptsPdf"":""true""}}"
                                               
      templateContent = "<#list reports as report>" & vbCrLf & "    ${report.reportID};<#t>" & vbCrLf & "</#list>"

ElseIf requestType = "policyList" Then

    ' Get list of Policies
    Debug.Print "Requesting list of Policies from Expensify..."
    postdata = "requestJobDescription={" _
                                        & """type"":""get""," _
                                        & """credentials"":{" _
                                        & """partnerUserID"":""" & userID & """," _
                                       & """partnerUserSecret"":""" & userSecret & """}," _
                                       & """inputSettings"":{" _
                                            & """type"":""policyList""}}"

ElseIf requestType = "policyGet" Then
    Debug.Print "Requesting Policy info from Expensify..."
    ' Get the Policy Categories and reportFields
    fileType = "csv"
    Dim fields As String
    fields = "[""categories"",""reportFields"",""tags""]"
    postdata = "requestJobDescription={" _
                                        & """type"":""get""," _
                                        & """credentials"":{" _
                                        & """partnerUserID"":""" & userID & """," _
                                       & """partnerUserSecret"":""" & userSecret & """}," _
                                       & """inputSettings"":{" _
                                            & """type"":""policy""," _
                                            & """fields"": " & fields & "," _
                                            & """policyIDList"": [""" & policyID & """]}}"
                                            
 ElseIf requestType = "policyUpdate" Then
    Debug.Print "Updating Policy info in Expensify..."
    
    ' Get the Policy Categories and reportFields
    fileType = "csv"
    postdata = "requestJobDescription=" & "{""type"":""update""," _
                                        & """credentials"":{" _
                                        & """partnerUserID"":""" & userID & """," _
                                       & """partnerUserSecret"":""" & userSecret & """}," _
                                       & """inputSettings"":{" _
                                            & """type"":""policy"",""policyID"": """ & policyID & """}," _
                                            & """categories"":{""action"": ""replace"",""data"": [" _
                                            & "{""name"": ""Airfare"",""enabled"": true}," _
                                            & "{""name"": ""Auto/Fuel"",""enabled"": true}," _
                                            & "{""name"": ""Lodging"",""enabled"": true}," _
                                            & "{""name"": ""Meals"",""enabled"": true}," _
                                            & "{""name"": ""Mileage"",""enabled"": true}," _
                                            & "{""name"": ""Other"",""enabled"": true}," _
                                            & "{""name"": ""Park/Tolls"",""enabled"": true}," _
                                            & "{""name"": ""Rental"",""enabled"": true}," _
                                            & "{""name"": ""Taxi"",""enabled"": true}," _
                                            & "{""name"": ""Train"",""enabled"": true}]}," _
                                            & """reportFields"":{""action"": ""replace"",""data"": [" _
                                            & "{""name"": ""Customer"",""type"":""text"",""setRequired"": false}," _
                                            & "{""name"": ""Week"",""type"":""text"",""setRequired"": false}," _
                                            & "{""name"": ""Reason for Trip"",""type"":""dropdown"",""setRequired"": false,""values"":[""Applications Training"",""ESL Owned System"",""Free of Charge"",""Home Office"",""Installation"",""Misc"",""Mobile Bills"",""Service Contract"",""Staff Training"",""Time and Material"",""Warranty""]}," _
                                            & "{""name"": ""System Type"",""type"":""dropdown"",""setRequired"": false,""values"":[""Artifact"",""ESL193fx"",""ESL193HE"",""ESL193UC"",""ESL213"",""ESLfemto"",""imageBIO"",""imageGEO"",""LaserSC"",""Lumen"",""MicroMill"",""MIR10"",""Other"",""n / a""]}," _
                                            & "{""name"": ""Serial Number"",""type"":""text"",""setRequired"": false}," _
                                            & "{""name"": ""Company"",""type"":""dropdown"",""setRequired"": false,""values"":[""Elemental Scientific Glassblowing"",""Elemental Scientific Inc."",""Elemental Scientific Instruments Ltd."",""Elemental Scientific Lasers LLC""],""defaultValue"":""Elemental Scientific Lasers LLC""}" _
                                            & "]}}"

End If

' URL encode the template content
templateContentURL = URLEncode(templateContent)

' Create the HTTP object
Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

' Open a POST request
httpRequest.Open "POST", url, False

' Set the request headers
httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"  ' Set as per Expensify API

fullSend = postdata & "&" & "template=" & templateContentURL

' Send the Request
httpRequest.Send fullSend
  
If httpRequest.Status = 200 Then
  Dim responseText As String
  expensifyRequestReturn = httpRequest.responseText

    
  If requestType = "csv" Then
    expensifyAPIRequest = expensifyRequestReturn
    Debug.Print "Request successful! File Name: " & expensifyRequestReturn
    
  ElseIf requestType = "PDF" Then
    expensifyAPIRequest = expensifyRequestReturn
    Debug.Print "Request successful! File Name: " & expensifyRequestReturn
    
  ElseIf requestType = "policyList" Then
    expensifyAPIRequest = expensifyRequestReturn
    Debug.Print "Request successful! Policy list is: " & expensifyRequestReturn
    
  ElseIf requestType = "policyGet" Then
    expensifyAPIRequest = expensifyRequestReturn
    Debug.Print "Request successful! Policy info is: " & expensifyRequestReturn
  
  ElseIf requestType = "policyUpdate" Then
  expensifyAPIRequest = expensifyRequestReturn
   Debug.Print httpRequest.statusText
   Debug.Print "Request successful! Policy info has been updated!"
  
  ElseIf requestType = "checkReimbursed" Then
    expensifyAPIRequest = expensifyRequestReturn
  Debug.Print "Request successful! Reimbursed Policy list is: " & expensifyRequestReturn
  
  ElseIf requestType = "combinedReports" Then
    expensifyAPIRequest = expensifyRequestReturn
  Debug.Print "Request successful! Combined list filename is: " & expensifyRequestReturn
  
  End If
  
Else
  Debug.Print "Error: " & httpRequest.statusText
End If

Set httpRequest = Nothing

End Function
Function expensifyAPIdownload(fileName As String, url As String, userID As String, userSecret As String, fileType As String, requestType As String, Optional filePath As String) As String

  Dim postdata As String
  Dim httpRequest As Object
  Dim templateContent As String
  Dim data() As Byte
  Dim fso As Object
  Dim templateContentURL As String
  Dim fullSend As String
  Dim response As String

  ' Set the URL and data
  postdata = "requestJobDescription={" _
                                    & """type"":""download""," _
                                    & """credentials"":{" _
                                    & """partnerUserID"":""" & userID & """," _
                                    & """partnerUserSecret"":""" & userSecret & """}," _
                                    & """fileName"":""" & fileName & """,""fileSystem"":""integrationServer""}}"
      
  'templateContent = "<#list reports as report>" & vbCrLf & "    ${report.reportID},<#t>" & vbCrLf & "</#list>"

  ' URL encode the template content
  templateContentURL = URLEncode(templateContent)

  ' Create the HTTP object
  Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

  ' Open a POST request
  httpRequest.Open "POST", url, False  ' True for asynchronous request

  ' Set the request headers
  httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"  ' Set for JSON data

  'fullSend = postdata & "&" & "template=" & templateContentURL
  fullSend = postdata


  ' Set the data to send
  httpRequest.Send fullSend

  If httpRequest.Status = 200 Then
    If requestType = "PDF" Then ' Or requestType = "csv" Then
    Application.StatusBar = "Requesting PDF from Expensify..."
    ' Get the response body as a byte array
    data = httpRequest.responseBody
    'Debug.Print httpRequest.responseBody
    'Debug.Print httpRequest.responseText
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Open file for binary access
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Binary As #fileNum

    ' Write byte array to file
    Put #fileNum, , data

    Close #fileNum

    Debug.Print "Expensify " & fileType & " downloaded successfully to: " & filePath
    
    Else ' Get the httpResponse for the reports and place it in "response"
      expensifyAPIdownload = httpRequest.responseText
      
    End If
 
  Else
    Debug.Print "Error downloading file: " & httpRequest.statusText
  End If

  Set httpRequest = Nothing
  Set fso = Nothing
End Function

