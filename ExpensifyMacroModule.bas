Attribute VB_Name = "ExpensifyMacroModule"
Option Explicit

' Expensify Conversion Macro
' v3.0 (240723) - implemented auto creation of directory structure
' v3.1 (240726) - implementing auto get of report pdf and placing in directory
' v4.0 (240727) - Implement autocreation of Email (Requires Outlook to be configured, not New Outlook!)
' v5.0 (240729) - Clean up code, isolate fuctions into separate subs
' v5.1 (240731) - Store settings in registry
' v5.2 (240731) - Add options to Tracking/Macro sheet
' v5.3 (240801) - Expand Options
' v6.0 (240805) - Pull in trip type to estimate bonus
' v6.1 (240920) - Update to how emails are generated
' v6.2 (240921) - Updates to store mileage info in new ESL Template (cheers Roblet!)
' v6.3 (250124) - Prevent privicy levels stopping xe.com lookup
' v6.4 (250124) - Stop OneDrive at start of Macro, then restart when finished
' v7.0 (250220) - Major refactoring and clean up of code
' v7.1 (250417) - Add tolerance factor to check for difference between ESL Report and Expensify
' v7.2 (250417) - Automatically update log and download report based on difference if already in Log
' v8.0 (250418) - Design new Settings Page - move variables into hidden worksheet
' v8.1 (250419) - Calculate fees based on Settings not Expensify Report
' v8.2 (250419) - If expense not categorised send user to Expensify Report or assign to Other
' v8.3 (250419) - Include line in Log to show that report needs updating to process


' Todo
' - check settings have been applied to Expensify correctly after initial onboarding
' - prevent crash if no reimbursed reports
' - only create one email option - put details and file names in to an array, then loop through array to populate email
' - create settings form
' - Create new sheet for new year
' - End of year overview report for Chris
' - Dashboard on sheet 1
' - figure out upgrade process - copy module in to current workbook
' - track total miles (apr to apr)

Sub ConvertExpensifyExpenses()
Dim requestType As String
Dim reportsFileResult As String
Dim reportsResult As String
Dim reportID As String
Dim reportName As String
Dim weekDirectory As String
Dim filePath As String
Dim fileType As String
Dim csvWb As Workbook
Dim csvPath As String
Dim reportsArray() As String
Dim i As Integer, l As Integer, m As Integer, n As Integer
Dim downloadresponse As String
Dim dataArr() As Variant
Dim numRows As Long, numCols As Long
Dim lines() As String
Dim wkNr As String
Dim customer As String
Dim name As String
Dim reasonForTrip As String
Dim serialNumber As String
Dim systemType As String
Dim saveDirectory As String
Dim expectedValue As Double
Dim logRow As Integer
Dim logStart As Integer
Dim wbLog As Workbook, wbOutput As Workbook
Dim wsLog As Worksheet, wsOutput As Worksheet
Dim feesTotal As Double
Dim mileagePayout As Double
Dim checkCell As Double
Dim repCurr As String
Dim expsIntChoice As String
Dim filePicker As Object
Dim fldrPicker As Object
Dim startTime As Double
Dim emptyIdentifier As String
Dim lowerLimit As Double, upperLimit As Double, withinPercent As Double
Dim loggingRange As Range
Dim foundCell As Range
Dim lastRow As Long
Dim columns() As String
Dim submittedDate As Date
Dim createdDate As Date
Dim numReports As Integer
Dim reportStatus As String
Dim mappedCategory As String
Dim foundRow As Range
Dim previousYear As Integer
Dim reportTotal As Double
Dim ESLTotal As Double ' Value from log file Column G (ESLTotal)
Dim searchLastRow As Long    ' Last row in log for searching
Dim skipProcessing As Boolean ' Flag to skip processing current archived report
Dim allowedDifference As Double ' Calculated difference between ESL and Expensify Totals based on set tolerance
Dim numReimbursedReports As Integer ' Number of Reimbursed Reports in array
Dim numArchivedReports As Integer ' Number of Submistted Reports in array
Dim numProcessedReports As Integer ' Current Report Number processing
Dim conversionSuccessful As Boolean


 If expensesDir = "" Then

      Debug.Print "Settings have not been loaded into memory. Trying to load now..."
      modSettings.LoadSettingsIntoMemory
    ' Re-check after attempting load
      If expensesDir = "" Then
        MsgBox "Failed to load settings. Please configure settings via the form.", vbCritical
             Exit Sub
         End If
     End If

  'Stop OneDrive
  If stopOneDrive = "TRUE" Then
    StopOneDriveSync
    Application.StatusBar = "Stopping OneDrive!"
    Debug.Print "Stopping OneDrive!"
    startTime = Timer
    Do While Timer < startTime + 2
      DoEvents ' Yield processing time to other applications
    Loop
  End If

  'Get Expensify API login details (https://www.expensify.com/tools/integrations/)
  Const url As String = "https://integrations.expensify.com/Integration-Server/ExpensifyIntegrations"

  
  ' Check if the Policy Configured value is set, otherwise configure policy
  If policyConfigured = False Then
    requestType = "policyUpdate"
    Debug.Print "Policy not configured via API yet. Performing initial setup..."
    Application.StatusBar = "Updating Policy info..."
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    Debug.Print reportsResult
    If CStr(reportsResult) = "{""responseCode"":200}" Then
            Debug.Print "Policy Update API call successful."

            ' --- Set the flag NOW that the update succeeded ---
            If modSettings.SaveSpecificSetting(modSettings.POLICY_CONFIGURED_CELL, True) Then
                modSettings.policyConfigured = True ' Update memory variable too!
                Debug.Print "Policy Configured flag set to True."
            End If
    End If
  End If
  
requestType = "combinedReports"

Start: ' Jump to here to re run after changing request type

  Set wbLog = Workbooks.Open(loggingFile)
  Set wsLog = wbLog.Sheets("Expense Logging")

  logStart = 10 'starting Row of log file
  logRow = 10 'Row to write data to
  
' manually set a request type for debugging

'requestType = "csv"
'requestType = "PDF"
'requestType = "policyList"
'requestType = "policyGet"
'requestType = "policyUpdate"
'requestType = "combinedReports"

If requestType = "csv" Then
    fileType = "csv"
    ' Create file path
    filePath = weekDirectory & reportName & "." & fileType
    
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    
    ' Import csv into excel
    csvPath = filePath
    Set csvWb = Workbooks.Open(fileName:=csvPath)
    
    Call ConvertExpensifyExpenses

ElseIf requestType = "PDF" Then
    fileType = "PDF"
    filePath = weekDirectory & reportName & "." & fileType
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    downloadresponse = expensifyAPIdownload(reportsResult, url, userID, userSecret, fileType, requestType, filePath)
    Debug.Print downloadresponse
    
ElseIf requestType = "policyList" Then
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    'Debug.Print reportsResult

ElseIf requestType = "policyGet" Then
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    'Debug.Print reportsResult
    
ElseIf requestType = "combinedReports" Then
    Application.StatusBar = "Requesting list of closed and reimbursed Reports from Expensify..."
    Debug.Print "Requesting list of closed and reimbursed Reports from Expensify..."
    
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    requestType = "csv"
    reportsResult = expensifyAPIdownload(reportsResult, url, userID, userSecret, fileType, requestType, filePath)
    
    ' Split the string into rows based on newlines
    lines = Split(reportsResult, vbNewLine)
    
    ' Determine the number of rows and columns
    numRows = UBound(lines) + 1
    numCols = UBound(Split(lines(0), ";")) + 1

    ' Resize the data array
    ReDim reportsArray(1 To numRows, 1 To numCols)

    ' Populate the data array
    For l = 1 To numRows
        columns = Split(lines(l - 1), ";")
        For m = 1 To numCols
            reportsArray(l, m) = Replace(columns(m - 1), """", "") ' Remove double quotes from all fields
        Next m
    Next l
    
    numReports = numRows
    
    ' Calculate Number of reports from the array
    numReimbursedReports = 0
    numArchivedReports = 0
    numProcessedReports = 1
    
    For i = LBound(reportsArray, 1) To UBound(reportsArray, 1)
      If reportsArray(i, 2) = "Reimbursed" Then
        numReimbursedReports = numReimbursedReports + 1
      End If
      
      If reportsArray(i, 2) = "Archived" Then
        numArchivedReports = numArchivedReports + 1
      End If
    Next i

    ' Define the range of your table (adjust columns as needed)

    Set wbLog = Workbooks.Open(loggingFile)
    Set wsLog = wbLog.Sheets("Expense Logging")
    
    With wbLog.Sheets("Expense Logging")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set loggingRange = .Range("A10:M" & lastRow) ' Adjust columns as needed
    End With
    
        ' process each report from the array
    For i = LBound(reportsArray, 1) To UBound(reportsArray, 1)
        skipProcessing = False
        requestType = "csv"
        fileType = "csv"
        reportID = reportsArray(i, 1)
        reportStatus = reportsArray(i, 2)
        submittedDate = reportsArray(i, 5)
        createdDate = reportsArray(i, 6)
        reportTotal = reportsArray(i, 7)

      If reportsArray(i, 2) = "Reimbursed" Then
        reportID = reportsArray(i, 1)

        ' Find the report ID in the table
        Set foundCell = loggingRange.Find(what:=reportID, LookIn:=xlValues, LookAt:=xlWhole)

        If Not foundCell Is Nothing And foundCell.Offset(0, 10).Value = "Submitted" Then
            ' Update the status to "Reimbursed"
            Application.StatusBar = "Setting Report " & reportID & " as reimbursed!"
            Debug.Print "Setting Report " & reportID & " as reimbursed!"
            foundCell.Offset(0, 10).Value = "Reimbursed"
        End If

      ElseIf reportsArray(i, 2) = "Archived" Then
      
        Debug.Print "Checking log file for Submitted Report ID: " & reportID; ""
        Application.StatusBar = "Checking log for Submitted Report " & reportID & "..."
        
      
        ' Find the last row with data (assuming headers are in row 1)
        lastRow = wsLog.Cells(logStart, 1).End(xlDown).Row
        
        ' Look for existing reportID in column B (assuming reportID is in column B)
        Set foundRow = wsLog.Range("B" & logStart & ":B" & lastRow).Find(reportID, LookIn:=xlValues, SearchOrder:=xlByRows)
        
        If Not foundRow Is Nothing Then
            ' Report ID FOUND in the log file
            Debug.Print "   FOUND Report ID '" & reportID & "' in log row " & foundRow.Row & "."
            If foundRow.Offset(0, 10).Value = "Reimbursed" Then
              foundRow.Offset(0, 10).Value = "Submitted"
            End If
            ESLTotal = 0 ' Default
            On Error Resume Next ' Handle non-numeric value in Col G
            ESLTotal = CDbl(foundRow.Offset(0, 5).Value) ' Column G is 5 cols offset from B
            If Err.Number <> 0 Then
                Debug.Print "      Warning: Could not convert value '" & foundRow.Offset(0, 5).Value & "' in log cell F" & foundRow.Row & " to a number."
                Err.Clear
                ESLTotal = -888888.88 ' Use a value guaranteed not to match
            End If
            On Error GoTo 0

            ' Compare totals
            allowedDifference = Abs(reportTotal * tolerance / 100)
            
            If Abs(ESLTotal - reportTotal) <= allowedDifference Then
                ' MATCH FOUND! Report ID exists and total matches. Skip processing.
                Debug.Print "      MATCH: ESL Total (" & ESLTotal & ") matches Expensify Total (" & reportTotal & "). Skipping processing."
                Application.StatusBar = "Submitted Report " & reportID & " already verified in log. Skipping."

                skipProcessing = True ' *** Set flag to skip ***
            Else
                ' MISMATCH: Report ID found, but total is different. Continue processing.
                Debug.Print "      MISMATCH: ESL Total (" & ESLTotal & ") differs from Expensify Total (" & reportTotal & "). Will re-process."
                Application.StatusBar = "Submitted Report " & reportID & " found in log, but TOTAL MISMATCH. Re-processing."
                ' Optional: Highlight the row in the log
                ' foundRow.EntireRow.Interior.Color = vbYellow
            End If
        Else
            ' Report ID NOT FOUND in the log file. Continue processing.
            Debug.Print "   NOT FOUND: Submitted Report ID '" & reportID & "' in log file (Rows " & logStart & "-" & searchLastRow & "). Proceeding with processing."
            Application.StatusBar = "Submitted Report " & reportID & " not found in log. Processing..."
        End If
      
    If Not skipProcessing Then
    
        Application.StatusBar = "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Requesting expense information for Report " & reportID & " - Week " & reportsArray(i, 4) & " - " & reportsArray(i, 3) & " from Expensify..."
        Debug.Print "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Requesting expense information for Report " & reportID & " - Week " & reportsArray(i, 4) & " - " & reportsArray(i, 3) & " from Expensify..."
        
        
        'Get report filename
        reportsFileResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
        
        'download csv info into an array
        Debug.Print "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Grabbing data from Expensify..."
        Application.StatusBar = "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Grabbing data from Expensify..."
                
        reportsResult = expensifyAPIdownload(reportsFileResult, url, userID, userSecret, fileType, requestType)
        
        ' split info in to rows
        lines = Split(reportsResult, vbNewLine)

        ' Determine the number of rows and columns
        numRows = UBound(lines) + 1
        numCols = UBound(Split(lines(0), ";")) + 1

        ' Resize the data array
        ReDim dataArr(1 To numRows, 1 To numCols)
    
        ' Populate the data array
        For l = 1 To numRows
            columns = Split(lines(l - 1), ";")
            For m = 1 To numCols
                dataArr(l, m) = Replace(columns(m - 1), """", "") ' Remove double quotes from all fields
            Next m
        Next l
        
        ' Produce ESL Report
        Debug.Print "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Producing ESL Expense Report..."
        Application.StatusBar = "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Producing ESL Expense Report..."
        Application.ScreenUpdating = False
        conversionSuccessful = expensesConversionDataArr(reportsResult, dataArr, wkNr, customer, name, ESLExpenseTemp, expensesDir, saveDirectory, reportName, serialNumber, systemType, reasonForTrip, expectedValue, repCurr, feesTotal, mileagePayout, checkCell, mappedCategory)
        
        ' Save totals to log file
        Application.StatusBar = "Saving expense data to log file..."
        Debug.Print "Saving expense data to log file..."
        ' Check if reportID already exists and update that row only
        
        ' Find the last row with data (assuming headers are in row 1)
        lastRow = wsLog.Cells(logStart, 1).End(xlDown).Row
        
        If wsLog.Range("A" & logStart).Value <> "" Then ' Check if first cell of log is empty
        
        ' Look for existing reportID in column B (assuming reportID is in column B)
            Set foundRow = wsLog.Range("B" & logStart & ":B" & lastRow).Find(reportID, LookIn:=xlValues, SearchOrder:=xlByRows)
        
            If foundRow Is Nothing Then  ' reportID not found, insert new row
                Debug.Print "No existing ReportID found, inserting new Row"
                wsLog.Rows(logStart).Insert xlShiftDown
            Else  ' reportID found, update existing row
                Debug.Print "Existing reportID found, updating that row!"
                logRow = foundRow.Row
            End If
        End If

        ' Write Expensify and output sheet totals
        wsLog.Range("A" & logRow).Value = submittedDate
        wsLog.Range("B" & logRow).Value = reportID
        wsLog.Range("C" & logRow).Value = wkNr
        wsLog.Range("D" & logRow).Value = mappedCategory
        wsLog.Range("E" & logRow).Value = customer
        wsLog.Range("F" & logRow).Value = expectedValue
        
        If feesTotal <> 0 Then
          wsLog.Range("I" & logRow).Value = feesTotal
        End If
        wsLog.Range("K" & logRow).Value = repCurr
        wsLog.Range("C3").Value = repCurr
        If mileagePayout <> 0 Then
          wsLog.Range("J" & logRow).Value = mileagePayout
        End If


       If conversionSuccessful Then
        wsLog.Range("L" & logRow).Value = "Submitted"
        
        'Close the Workbook
        If closeReport = "TRUE" Then
            Workbooks(reportName & ".xlsx").Close SaveChanges:=False
        End If
        Application.ScreenUpdating = True
        
        ' Request pdf for report
        Debug.Print "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Requesting PDF and getting filename from Expensify..."
        Application.ScreenUpdating = True
        Application.StatusBar = "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Requesting PDF and getting filename from Expensify..."
        requestType = "PDF"
        fileType = "pdf"
        filePath = saveDirectory & reportName & "." & fileType
        reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)

        ' download PDF
        Debug.Print "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Downloading PDF..."
        Application.ScreenUpdating = True
        Application.StatusBar = "Processing report " & numProcessedReports & " of " & numArchivedReports & ": Downloading PDF..."
        downloadresponse = expensifyAPIdownload(reportsResult, url, userID, userSecret, fileType, requestType, filePath)
        
        wsLog.Hyperlinks.Add Range("M" & logRow), Address:="file:///" & saveDirectory & reportName & ".xlsx", TextToDisplay:="Link"
        wsLog.Hyperlinks.Add Range("N" & logRow), Address:="file:///" & saveDirectory & reportName & ".pdf", TextToDisplay:="Link"
        
        ' Calculate difference
        wsLog.Range("G" & logRow).Value = checkCell
        wsLog.Range("H" & logRow).Value = checkCell - expectedValue
        
        ' Calculate allowed range based on tolerance
      
        lowerLimit = expectedValue * (1 - (tolerance / 100)) ' 5% less than expected value
        upperLimit = expectedValue * (1 + (tolerance / 100)) ' 5% more than expected value
      
        If checkCell < lowerLimit Or checkCell > upperLimit Then ' Check if within tolerance range
          MsgBox "The ESL Total (" & checkCell & ") is not within " & tolerance & "% of the Expensify Report Total (" & expectedValue & "). Check you correctly categorised your Expenses!", vbExclamation
        End If
        
        'Send Email if required
        
        If createEmail = "Individual" Then
        Application.StatusBar = "Creating Email..."
        Debug.Print "Creating Email..."
        Call Send_Email(wkNr, customer, reasonForTrip, systemType, saveDirectory, reportName, serialNumber, name)
        End If 'Email If
      Else
        ' --- FAILURE or EARLY EXIT ---
        ' The function did NOT complete successfully. The workbook was likely closed
        ' without being saved (in the Exit Function path), or a critical error occurred.
        ' Skip the steps that rely on the workbook existing/being processed.
        Debug.Print "Conversion was not successful or was exited early. Skipping subsequent steps for this report."
        wsLog.Range("L" & logRow).Value = "Incomplete"
      End If
        
      With wsLog.ListObjects("Table1").Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsLog.Range("Table1[Date Submitted]"), Order:=2
            .Header = xlYes
            .Apply
      End With
        
        Workbooks(loggingFilenameOnly).RefreshAll

        'Save the logbook
        Workbooks(loggingFilenameOnly).Save
        
        ' increment counter for reports processed
        numProcessedReports = numProcessedReports + 1
    End If
  End If
  Next i
  
    'Restart Onedrive
    If stopOneDrive = "TRUE" Then
      Application.StatusBar = "Restarting OneDrive!"
      Debug.Print "Restarting OneDrive!"
      StartOneDriveSync
      startTime = Timer
      Do While Timer < startTime + 1

        DoEvents ' Yield processing time to other applications
      Loop
    End If
  
  'Show completed and wait 5 seconds until clear
  Application.StatusBar = "Complete!"
  Debug.Print "Complete!"
  startTime = Timer
  Do While Timer < startTime + 5
      DoEvents ' Yield processing time to other applications
  Loop

Application.StatusBar = ""


End If

End Sub
