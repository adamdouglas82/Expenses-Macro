Attribute VB_Name = "ExpensesConversion"
Function expensesConversionDataArr(reportsResult, dataArr, wkNr, customer, name, ESLExpenseTemp, expensesDir, ByRef saveDirectory, ByRef reportName, ByRef serialNumber, ByRef systemType, ByRef reasonForTrip, ByRef expectedValue, ByRef repCurr, ByRef feesTotal, ByRef mileagePayout, ByRef checkCell, ByRef mappedCategory)
' Expensify Conversion Macro
' v3.0 (240723) - implemented auto creation of directory structure
' v3.1 (240726) - implementing auto get of report pdf and placing in directory
' v4.0 (240727) - Implement autocreation of Email (Requires Outlook to be configured, not New Outlook!)
' v5.0 (240731) - Clean up code, isolate fuctions into separate subs

  Dim wbNew As Workbook
  Dim wsInput As Worksheet, wsOutput, wsMileage As Worksheet
  Dim dataRow As Long, outputRow As Long, mileageRow As Long  ' Separate variables for row tracking
  Dim iCol As Long
  Dim result As Variant
  Dim lookupValue As String
  Dim mileageTotal As Double  ' Variable to store total mileage
  Dim colHeaders As Object  ' Declare a dictionary object
  'Dim mappedCategory As String
  Dim minDate As Variant
  Dim maxDate As Variant
  Dim currentWorkbook As Variant
  Dim oneDrivePaths() As Variant
  Dim commonPath As String
  Dim firstExpenseYear As Integer
  Dim yearDirectory As String
  Dim reportID As String
  Dim policyID As String
  Dim PDFFilename As String
  Dim wsLog As Worksheet
  Dim logRow As Long
  Dim expensifyTotal As Double
  Dim outputTotal As Double
  Dim difference As Double
  Dim lines() As String
  Dim fields() As String
  Dim i As Long, j As Long, csvrow As Long, csvcol As Long
  Dim headerValue As String
  Dim env As Variant
  Dim lastBackslash As String
  Dim directoryPath As String
  Dim expenseType As String
  Dim mileageRate As String
  Dim reportcurrencyresult As String
  Dim col As Integer
  Dim loggingPath As String
  Dim wbSource As Object
  Dim currentField As String, inQuotes As Boolean
  
  ' Declare variables for column indexes
  Dim subCol As Long
  Dim dateCol As Long
  Dim descCol As Long
  Dim notesCol As Long
  Dim catCol As Long
  Dim expamtCol As Long
  Dim expCurrCol As Long
  Dim exrateCol As Long
  Dim conamtCol As Long
  Dim repCurrCol As Long
  Dim mileCol As Long
  Dim milerateCol As Long
  Dim mileunitCol As Long
  Dim custCol As Long
  Dim weekCol As Long
  Dim reasonCol As Variant
  Dim sysCol As Long
  Dim compCol As Long
  Dim serCol As Long
  Dim TotalCol As Long
  Dim feesCol As Long
  Dim policyCol As Long
  Dim reportIDCol As Long
  
  ' Create the regular expression object.
   Set regex = CreateObject("VBScript.RegExp")

   With regex
       .Global = True        ' We only expect one match, but this is good practice.
       .IgnoreCase = True   ' Case-insensitive (not strictly needed here)
       '.Pattern = "([\d.]+)\s*@\s*([\d.]+)"  ' The simplified regex pattern
       ' Regex Pattern: Handles optional quotes AND optional currency symbols (£, $, u20ac)
      .Pattern = """?([\d.]+)\s*[a-zA-Z]+\s*@\s*(?:£|\$|u20ac)?([\d.]+)\s*/\s*[a-zA-Z]+""?"
   End With

  ' Create a dictionary to store column headers and their indexes
  Set colHeaders = CreateObject("Scripting.Dictionary")
  
  ' Loop through the first row (headers)
  For iCol = 1 To UBound(dataArr, 2)  ' Use dataArr for column count
    ' Get the value from the current column in the first row (dataArr)
    headerValue = dataArr(1, iCol)

    ' Add header (key) and index (value) to dictionary (if not empty)
    If headerValue <> "" Then  ' Check if header is not empty
      colHeaders.Add headerValue, iCol
    End If
  Next iCol
  
  'define any needed headers
  subCol = colHeaders("Submitted by:")
  dateCol = colHeaders("Date")
  descCol = colHeaders("Description")
  notesCol = colHeaders("Meal Members/Notes")
  catCol = colHeaders("Category")
  expamtCol = colHeaders("Expense Amount")
  expCurrCol = colHeaders("Expense Currency")
  exrateCol = colHeaders("Exchange Rate")
  conamtCol = colHeaders("Converted Amount")
  repCurrCol = colHeaders("Report Currency")
  mileCol = colHeaders("Mileage")
  milerateCol = colHeaders("Mileage Rate")
  mileunitCol = colHeaders("Mileage Unit")
  custCol = colHeaders("Customer")
  weekCol = colHeaders("Week")
  reasonCol = colHeaders("Reason for Trip")
  sysCol = colHeaders("System Type")
  compCol = colHeaders("Company")
  serCol = colHeaders("Serial Number")
  TotalCol = colHeaders("Report Total")
  feesCol = colHeaders("Fees")
  policyCol = colHeaders("Policy")
  reportIDCol = colHeaders("Report ID")
  
    ' Get week number from array
  wkNr = dataArr(2, weekCol)
  wkNr = Trim(wkNr)
  
  ' Get customer from array
  customer = dataArr(2, custCol)
  customer = Trim(customer)
  
    ' Get Submitter Name From Array
  name = dataArr(2, subCol)
  name = Trim(name)
    
  ' Get serial number from array
  serialNumber = dataArr(2, serCol)
  serialNumber = Trim(serialNumber)

  ' Get system type
  systemType = dataArr(2, sysCol)
  systemType = Trim(systemType)
  
  ' Get reason for Trip
  reasonForTrip = dataArr(2, reasonCol)
  reasonForTrip = Trim(reasonForTrip)
  
  'Get expected total from array
  expectedValue = dataArr(2, TotalCol)
  
  'Get report Currency from array
  repCurr = dataArr(2, repCurrCol)
  
  ' map Reason for Work to Category of Work
  Select Case reasonForTrip
    Case "Install", "Installation"  ' Handle variations
        mappedCategory = "Install"
    Case "Service Contract", "Time and Material", "Warranty", "Free of Charge", "Applications", "ESL Owned"
        mappedCategory = "Service"
    Case "Staff Training", "Misc"
        mappedCategory = "Other"
    Case Else  ' Handle unmapped categories / spelling mistakes
        mappedCategory = "Other"
        
  End Select

  'Get Policy ID from array
  policyID = dataArr(2, policyCol)
  
  'Get Report ID from Array
  reportID = dataArr(2, reportIDCol)


  'find last backslash in ESLExpenseTemp
  lastBackslash = InStrRev(ESLExpenseTemp, "\")

  ' Extract the directory path by taking everything up to the last backslash
  directoryPath = Left(ESLExpenseTemp, lastBackslash)
  
  ' Create new workbook from template
  Set wbNew = Workbooks.Add(template:=ESLExpenseTemp)
  wbNew.Queries.FastCombine = True ' Enable Fast Combine to bypass privacy checks
  Set wsOutput = wbNew.Sheets("Sheet1")
  Set wsMileage = wbNew.Sheets("Mileage")

  ' Copy Week Number, Name and Business to correct parts
  wsOutput.Range("J10").Value = systemType & " - " & reasonForTrip ' Write System Type - Reason
  wsOutput.Range("A10").Value = name ' Submitted by
  wsOutput.Range("I3").Value = dataArr(2, compCol) ' Company
  wsOutput.Range("B17").Value = repCurr ' Report Currency / Target Currency
  wsOutput.Range("G11").Value = mappedCategory ' Write mapped category to cell G11
  wsOutput.Range("D10").Value = serialNumber & " - " & customer ' Serial Number - Customer

  ' Select correct First Row for data in output sheet
  outputRow = 17
    
  mileageRow = 4 ' set starting row for mileage

  ' Loop through data and write to output sheet
  For dataRow = UBound(dataArr, 1) To 2 Step -1

  ' Check if expense type is "Mileage"

  'expenseType = dataArr(dataRow, catCol)
  'If expenseType = "Mileage" Then ' Process only Mileage Entries
    'mileageTotal = mileageTotal + CDbl(dataArr(dataRow, mileCol)) ' Calculate mileage total (assuming numeric value)
  expenseType = dataArr(dataRow, catCol)
  If expenseType = "Mileage" Then
  
      ' --- Check if mileage and rate are empty ---
      If IsEmpty(dataArr(dataRow, mileCol)) Or dataArr(dataRow, mileCol) = "" Or IsEmpty(dataArr(dataRow, milerateCol)) Or dataArr(dataRow, milerateCol) = "" Then
  
          ' --- Attempt to extract from descCol ---
          Set matches = regex.Execute(dataArr(dataRow, descCol))
  
          If matches.Count > 0 Then
              Set Match = matches(0)
  
              ' Extract captured groups (now just the numbers)
              mileageStr = Match.SubMatches(0)  ' First group: mileage
              rateStr = Match.SubMatches(1)     ' Second group: rate
  
              ' --- Convert to numeric values ---
              On Error Resume Next
              mileageValue = CDbl(mileageStr)
              rateValue = CDbl(rateStr)
              On Error GoTo 0
  
              ' --- Use the extracted values ---
              If Err.Number = 0 Then
                  wsMileage.Range("H" & mileageRow).Value = mileageValue
                  ' --- Output Mileage Rate if not GBP or USD (Corrected) ---
                  If dataArr(2, repCurrCol) <> "GBP" And dataArr(2, repCurrCol) <> "USD" Then
                      wsOutput.Range("P52").Value = rateValue
                  End If
              Else
                  Debug.Print "Error converting: " & dataArr(dataRow, descCol)
              End If
          Else
              Debug.Print "No mileage info in notes: " & dataArr(dataRow, descCol)
          End If
  
    Else
      
        mileageRate = dataArr(dataRow, milerateCol) / 100 ' Find mileage rate and divide by 100 due to expensify report bug
        ' Insert Miles
        wsMileage.Range("H" & mileageRow).Value = dataArr(dataRow, mileCol)
        
        ' Copy over Mileage rate if not GBP or USD
        If dataArr(2, repCurrCol) <> "GBP" And dataArr(2, repCurrCol) <> "USD" Then
          wsOutput.Range("P52").Value = mileageRate ' Output mileage rate if currency is not GBP or USD
        End If
    
    End If
    
    ' Find min and Max date
    If IsEmpty(minDate) Or minDate > dataArr(dataRow, dateCol) Then
          minDate = dataArr(dataRow, dateCol)
        End If
    If IsEmpty(maxDate) Or maxDate < dataArr(dataRow, dateCol) Then
          maxDate = dataArr(dataRow, dateCol)
        End If
    
    ' Insert Date
     wsMileage.Range("A" & mileageRow).Value = dataArr(dataRow, dateCol) ' Date
     
    ' Insert purpose
    wsMileage.Range("B" & mileageRow).Value = customer
    
    ' Pull in description and notes and place in to new mileage sheet
    wsMileage.Range("E" & mileageRow).Value = dataArr(dataRow, notesCol) ' notes section combined
    
    'increment mileage row counter
    mileageRow = mileageRow + 1
    
    
  Else
   
    ' Find min and Max date
    If IsEmpty(minDate) Or minDate > dataArr(dataRow, dateCol) Then
          minDate = dataArr(dataRow, dateCol)
        End If
    If IsEmpty(maxDate) Or maxDate < dataArr(dataRow, dateCol) Then
          maxDate = dataArr(dataRow, dateCol)
        End If

    ' Get correct currency format
    lookupValue = dataArr(dataRow, expCurrCol)
    reportcurrencyresult = Application.Index(Sheets("CurrencyList").Range("A1:A140"), Application.Match(lookupValue & "*", Sheets("CurrencyList").Range("A1:A140"), 0), 1)

    ' Write data to output sheet (for non-mileage entries)
    wsOutput.Range("A" & outputRow).Value = reportcurrencyresult ' Paste currency
    wsOutput.Range("C" & outputRow).Value = dataArr(dataRow, dateCol) ' Date
    wsOutput.Range("D" & outputRow).Value = dataArr(dataRow, descCol) ' Description
    wsOutput.Range("E" & outputRow).Value = dataArr(dataRow, notesCol) ' Meals / Notes

    ' Map expense type to category based on names in target sheet
    MatchFound = False
    expenseType = dataArr(dataRow, catCol) ' column containing the expense type
    For col = 6 To 14 ' Loop through category name cells (F16:N16)
      If expenseType = wsOutput.Range(Cells(16, col).Address) Then ' Check for matching expense type
        wsOutput.Range(Cells(outputRow, col).Address) = dataArr(dataRow, expamtCol) ' Write expense amount to category column
        MatchFound = True
        Exit For ' Exit loop after finding a match

      End If
      
    Next col
      
      ' write to the "Other" category to avoid missing
      If Not MatchFound Then
      
            ' Define the base URL structure - NOTE THE DOUBLE QUOTES needed for VBA strings!
            baseURL = "https://www.expensify.com/report?param={""pageReportID"":""" & reportID & """,""keepCollection"":true}"

            ' Prepare the message box
            msgText = "Expense " & dataArr(dataRow, dateCol) & " - " & dataArr(dataRow, descCol) & " - " & dataArr(dataRow, expamtCol) & dataArr(dataRow, expCurrCol) & " was not matched." & vbCrLf & vbCrLf & _
                      "Do you want to stop processing (ID: " & reportID & ") and open the Expensify report in your browser to categorize it now?" & vbCrLf & vbCrLf & _
                      "Clicking No will assign it to ""Other"""
            msgTitle = "Unmatched Category"
            msgButtons = vbQuestion + vbYesNo + vbDefaultButton1 ' Yes/No question, default to No

            ' Show the message box and get the user's response
            msgResult = MsgBox(msgText, msgButtons, msgTitle)

            ' Act based on the response
            If msgResult = vbYes Then
                ' User wants to open the link
                Debug.Print "User chose Yes. Opening URL for Report ID " & reportID & " and closing ESL Report"
                On Error Resume Next ' Handle errors opening the hyperlink (e.g., browser issues)
                wbNew.Close SaveChanges:=False ' Close wbNew, discard changes
                ThisWorkbook.FollowHyperlink Address:=baseURL
                If Err.Number <> 0 Then
                     MsgBox "Could not open the Expensify link in your browser." & vbCrLf & _
                            "URL: " & baseURL & vbCrLf & "Error: " & Err.Description, vbExclamation, "Hyperlink Error"
                     Err.Clear
                     ' Decide: Write to "Other" if link fails?
                     ' wsOutput.Cells(outputRow, 10).Value = expAmount
                End If
                On Error GoTo 0
                 ' Typically, if they open the link, you DON'T write to "Other" here.
                 ' The value is effectively deferred pending user action in Expensify.
                expensesConversionDataArr = False
                Exit Function
            Else
                ' User chose No - Decide what to do. Write to "Other" or skip?
                Debug.Print "User chose No. Expense type '" & expenseType & "' (Report ID: " & reportID & ") placed in Other."
                wsOutput.Cells(outputRow, 10).Value = expAmount ' Write to "Other" (Col J) if user clicks No
            End If
      End If

    
    ' Check if currency matches report currency and calculate fees
      If dataArr(dataRow, expCurrCol) <> dataArr(2, repCurrCol) And IsNumeric(defaultFees) And defaultFees <> 0 Then
        wsOutput.Range("S" & outputRow).Formula = "=ROUND(O" & outputRow & "*(" & defaultFees & "/100)*Q" & outputRow & ",2)"
      End If
    
    ' Increment Output Row Counter
    outputRow = outputRow + 1

  End If
   
  Next dataRow

  
  ' Put min and max date into cell on the output sheet
  wsOutput.Range("R10").Value = minDate
  wsOutput.Range("R11").Value = maxDate
   
  ' Refresh the data connection
  Application.ScreenUpdating = True
  Application.StatusBar = "Refreshing xe.com data..."
  Application.ScreenUpdating = False
  
  ActiveWorkbook.RefreshAll
  
  Application.ScreenUpdating = True
  Application.StatusBar = "Creating directories and saving report..."
  'Application.ScreenUpdating = False
  
  
   ' Calculate the Total fees
  feesTotal = WorksheetFunction.Sum(wsOutput.Range("S17:S47"))
  
  ' Store the mileage payback value
  mileagePayout = wsOutput.Range("T50")

  ' Check Total value in output sheet before saving actual cell reference)
  checkCell = wsOutput.Range("T51") - feesTotal  'cell reference for the total - additional card fees

  ' Calculate expected value from Expensify data
  expectedValue = dataArr(2, TotalCol)
  
  ' Create the year directory
  firstExpenseYear = Year(dataArr(UBound(dataArr, 1), dateCol))
  yearDirectory = expensesDir & firstExpenseYear
  
  If Dir(yearDirectory, vbDirectory) = "" Then
        MkDir yearDirectory
    End If

  ' Create the week directory
  saveDirectory = yearDirectory & "\" & "Week " & wkNr & " - " & customer & "\"
  
  If Dir(saveDirectory, vbDirectory) = "" Then
        MkDir saveDirectory
    End If
    
  ' Create report name
  reportName = name & " - Expense Report - Week " & wkNr & " - " & customer

  ' Directly Save, delete any present file
  If Dir(saveDirectory & reportName & ".xlsx") <> "" Then Kill (saveDirectory & reportName & ".xlsx")
  ActiveWorkbook.SaveAs fileName:=saveDirectory & reportName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
  expensesConversionDataArr = True
End Function
