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
' v7.0 (2502xx) - Major refactoring and clean up of code


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
Dim policyID As String
Dim downloadresponse As String
Dim dataArr() As Variant
Dim numRows As Long, numCols As Long
Dim expensesDir As String
Dim lines() As String
Dim wkNr As String
Dim customer As String
Dim name As String
Dim reasonForTrip As String
Dim serialNumber As String
Dim systemType As String
Dim ESLExpenseTemp As String
Dim userID As String
Dim userSecret As String
Dim saveDirectory As String
Dim loggingFile As String
Dim loggingPath As String
Dim expectedValue As Double
Dim logRow As Integer
Dim wbLog As Workbook, wbOutput As Workbook
Dim wsLog As Worksheet, wsOutput As Worksheet
Dim feesTotal As Double
Dim mileagePayout As Double
Dim checkCell As Double
Dim repCurr As String
Dim autoSend As String
Dim reset As String
Dim expsIntChoice As String
Dim filePicker As Object
Dim fldrPicker As Object
Dim autoRun As String
Dim reDownload As String
Dim createEmail As String
Dim startTime As Double
Dim emptyIdentifier As String
Dim lowerLimit As Double, upperLimit As Double, withinPercent As Double
Dim closeReport As String
Dim policyConfigured As String
Dim loggingRange As Range
Dim foundCell As Range
Dim lastRow As Long
Dim columns() As String
Dim submittedDate As Date
Dim createdDate As Date
Dim numReports As Integer
Dim reportStatus As String
Dim mappedCategory As String
Dim upgrade As String
Dim update As String
Dim useFees As String
Dim defaultFees As String
Dim foundRow As Range
Dim stopOneDrive As String
Dim previousYear As Integer
Dim reportTotal As Double

  ' Pull settings into variables
  autoRun = ActiveSheet.Range("K2").Value
  closeReport = ActiveSheet.Range("K3").Value
  createEmail = ActiveSheet.Range("K4").Value
  autoSend = ActiveSheet.Range("K5").Value
  reDownload = ActiveSheet.Range("M2").Value
  reset = ActiveSheet.Range("M3").Value
  update = ActiveSheet.Range("K6").Value
  upgrade = ActiveSheet.Range("M4").Value
  stopOneDrive = ActiveSheet.Range("K7").Value
  
  'Stop OneDrive
  If stopOneDrive = "Yes" Then
    StopOneDriveSync
    startTime = Timer
    Do While Timer < startTime + 5
      Application.StatusBar = "Stopping OneDrive!"
      DoEvents ' Yield processing time to other applications
    Loop
  End If

  'Get Expensify API login details (https://www.expensify.com/tools/integrations/)
  Const url As String = "https://integrations.expensify.com/Integration-Server/ExpensifyIntegrations"

  ' Attempt to retrieve the saved directory from the registry
  expensesDir = GetSetting("ExpensifyConversion", "Directories", "expensesDir", "")
  
  If expensesDir <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "Directories", "expensesDir"
    expensesDir = ""
  End If
  
  ' If no directory is found, prompt the user
  If expensesDir = "" Then
      Set fldrPicker = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4
      
      With fldrPicker
          .Title = "Select Expenses Directory..."
          .AllowMultiSelect = False
          
          If .Show = -1 Then ' User selected a folder
              expensesDir = .SelectedItems(1) & "\"
              
              ' Save the selected directory to the registry
              SaveSetting "ExpensifyConversion", "Directories", "expensesDir", expensesDir
          Else
              ' User canceled the dialog
              MsgBox "No save directory selected.", vbExclamation
              Exit Sub
          End If
      End With
  End If
  
  ' select ESL template location

  ' Attempt to retrieve the saved file path from the registry
  ESLExpenseTemp = GetSetting("ExpensifyConversion", "FileNames", "ESLexpenseTemp", "")
  
  If ESLExpenseTemp <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "FileNames", "ESLexpenseTemp"
    ESLExpenseTemp = ""
  End If

  ' If no file path is found, prompt the user
  If ESLExpenseTemp = "" Then
      Set filePicker = Application.FileDialog(3) ' msoFileDialogFilePicker = 3

      With filePicker
          .Title = "Select ESL Template File..."
          
          .AllowMultiSelect = False
          .InitialFileName = expensesDir

          If .Show = -1 Then ' User selected a file
              ESLExpenseTemp = Split(.SelectedItems(1), "\")(UBound(Split(.SelectedItems(1), "\")))

              ' Save the selected file path to the registry
              SaveSetting "ExpensifyConversion", "FileNames", "ESLexpenseTemp", ESLExpenseTemp
          Else
              ' User canceled the dialog
              MsgBox "No file selected.", vbExclamation
              Exit Sub
          End If
      End With
  End If

' select logging file location

  ' Attempt to retrieve the saved file path from the registry
  loggingFile = GetSetting("ExpensifyConversion", "FileNames", "loggingFile", "")

  If loggingFile <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "FileNames", "loggingFile"
    loggingFile = ""
  End If
  
  ' If no file path is found, prompt the user
  If loggingFile = "" Then
      Set filePicker = Application.FileDialog(3) ' msoFileDialogFilePicker = 3

      With filePicker
          .Title = "Select ESL logging File..."
          
          .AllowMultiSelect = False
          .InitialFileName = expensesDir

          If .Show = -1 Then ' User selected a file
              loggingFile = Split(.SelectedItems(1), "\")(UBound(Split(.SelectedItems(1), "\")))

              ' Save the selected file path to the registry
              SaveSetting "ExpensifyConversion", "FileNames", "loggingFile", loggingFile
          Else
              ' User canceled the dialog
              MsgBox "No file selected.", vbExclamation
              Exit Sub
          End If
      End With
  End If

  ' Check if the userID value exists in the registry
  userID = GetSetting("ExpensifyConversion", "UserData", "userID", "")

  If userID <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "UserData", "userID"
    userID = ""
  End If
  
  ' Value doesn't exist, prompt the user, then save setting
  If userID = "" Then
      expsIntChoice = MsgBox("Do you have your Expensify Integration UserID and Secret?" & vbNewLine & vbNewLine & "Pressing No will take you to the Expensify Integrations page to generate these.", vbYesNo, "Enter userID...")
      If expsIntChoice = vbNo Then
        ActiveWorkbook.FollowHyperlink Address:="https://www.expensify.com/tools/integrations/"
      End If
      userID = InputBox("Please enter Expensify Integrations User ID:", "UserID")
      SaveSetting "ExpensifyConversion", "UserData", "userID", userID
  End If
  
  ' Check if the userSecret value exists in the registry
  userSecret = GetSetting("ExpensifyConversion", "UserData", "userSecret", "")
  
  If userSecret <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "UserData", "userSecret"
    userSecret = ""
  End If
  
  ' Value doesn't exist, prompt the user, then save setting
  If userSecret = "" Then
      userSecret = InputBox("Please enter Expensify Integrations User Secret:", "User Secret")
      SaveSetting "ExpensifyConversion", "UserData", "userSecret", userSecret
  End If
  
  ' Check if the Policy ID value exists in the registry
  policyID = GetSetting("ExpensifyConversion", "UserData", "policyID", "")
  
  If policyID <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "UserData", "policyID"
    policyID = ""
  End If
  
  ' Value doesn't exist, prompt the user, then save setting
  If policyID = "" Then
      policyID = InputBox("Please enter Expensify Policy ID: " & vbNewLine & vbNewLine & "This can be found in the url when on Expensify when in Settings - Workspaces and your workspace page.", "Policy ID")
      SaveSetting "ExpensifyConversion", "UserData", "policyID", policyID
  End If
  
  ' Check if the use fees value exists in the registry
  useFees = GetSetting("ExpensifyConversion", "UserData", "useFees", "")
  defaultFees = GetSetting("ExpensifyConversion", "UserData", "defaultFees", "")
  
  If useFees <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "UserData", "useFees"
    DeleteSetting "ExpensifyConversion", "UserData", "defaultFees"
    useFees = ""
  End If
  
  ' Value doesn't exist, prompt the user, then save setting
  If useFees = "" Then
      useFees = MsgBox("Do you want to automatically calculate transactions fees based on a percentage?", vbYesNo, "Calculate Fees")
      If useFees = vbYes Then
        SaveSetting "ExpensifyConversion", "UserData", "useFees", "Yes"
        defaultFees = InputBox("Please enter default conversion percentage (%):", "defaultFees")
        SaveSetting "ExpensifyConversion", "UserData", "defaultFees", defaultFees
      Else
        SaveSetting "ExpensifyConversion", "UserData", "useFees", "No"
      End If
  End If
  
  ' Check if the Policy Configured value exists in the registry
  policyConfigured = GetSetting("ExpensifyConversion", "UserData", "policyConfigured", "")
  
  If policyConfigured <> "" And reset = "Yes" Then
    DeleteSetting "ExpensifyConversion", "UserData", "policyConfigured"
    policyConfigured = ""
  End If
  
  If policyConfigured = "" Then
      requestType = "policyUpdate"
      SaveSetting "ExpensifyConversion", "UserData", "policyConfigured", "Yes"
  Else

  End If
  
 If requestType = "" Then
      requestType = "combinedReports"
 End If
 
Start: ' Jump to here to re run after changing request type
   
  'Turn off Reset
  ActiveSheet.Range("M3").Value = "No"


' manually set a request type for debugging

'requestType = "csv"
'requestType = "PDF"
'requestType = "policyList"
'requestType = "policyGet"
'requestType = "policyUpdate"
'requestType = "checkReimbursed"
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
    
ElseIf requestType = "policyUpdate" Then
    Application.StatusBar = "Updating Policy info..."
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID, , useFees, defaultFees)
    requestType = "policyGet"
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    Call ConvertExpensifyExpenses
    'Debug.Print reportsResult

ElseIf requestType = "checkReimbursed" Then
    Application.StatusBar = "Checking for any new reimbursed Reports..."
    fileType = "csv"
    reportsFileResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
    reportsResult = expensifyAPIdownload(reportsFileResult, url, userID, userSecret, fileType, requestType)
    
    ' restore normal result string
    'reportsResult = Mid(reportsResult, Len(emptyIdentifier) + 1)
    
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
    
    ' Define the range of your table (adjust columns as needed)
    loggingPath = expensesDir & loggingFile

    Set wbLog = Workbooks.Open(loggingPath)
    Set wsLog = wbLog.Sheets("Expense Logging")
    
    With wbLog.Sheets("Expense Logging")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set loggingRange = .Range("A10:M" & lastRow) ' Adjust columns as needed
    End With

    ' Iterate through the report IDs
    For n = LBound(reportsArray, 1) To UBound(reportsArray, 1)
      If reportsArray(n, 2) = "Reimbursed" Then
        reportID = reportsArray(n, 1)

        ' Find the report ID in the table
        Set foundCell = loggingRange.Find(what:=reportID, LookIn:=xlValues, LookAt:=xlWhole)

        If Not foundCell Is Nothing Then
            ' Update the status to "Reimbursed"
            Application.StatusBar = "Setting Report " & foundCell.Offset(0, 2).Value & "as reimbursed!"
            foundCell.Offset(0, 10).Value = "Reimbursed"
        End If
      End If
        
    Next n
  startTime = Timer

  Do While Timer < startTime + 5
  Application.StatusBar = "Complete!"
      DoEvents ' Yield processing time to other applications
  Loop

Application.StatusBar = ""
    
    'Debug.Print reportsResult
    
ElseIf requestType = "combinedReports" Then
    Application.StatusBar = "Requesting list of closed and reimbursed Reports from Expensify..."
    Debug.Print "Requesting list of closed and reimbursed Reports from Expensify..."
    
    reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID, reDownload)
    
    'Turn off reDownload
    If reDownload = "All" Or reDownload = "Changed" Then
        ActiveSheet.Range("M2").Value = "Changed"
    End If
    requestType = "csv"
    reportsResult = expensifyAPIdownload(reportsResult, url, userID, userSecret, fileType, requestType, filePath)
    
    '' determine if empty then get reports list
    '
    'emptyIdentifier = "$empty$"
    'reportsResult = emptyIdentifier & Trim(reportsResult)
    '
    '' remove trailing comma or empty character from string
    'reportsResult = Left(reportsResult, Len(reportsResult) - 1)
    '
    'If Left(reportsResult, Len(emptyIdentifier) + 1) = emptyIdentifier Then
    '    MsgBox ("No closed reports found!" & vbNewLine & vbNewLine & "Please check for closed reports in Expensify." & vbNewLine & vbNewLine & "Checking for reimbursed reports...")
    '    requestType = "checkReimbursed"
    '    GoTo Start
    'End If
    '
    '' restore normal result string
    'reportsResult = Mid(reportsResult, Len(emptyIdentifier) + 1)
    
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

    ' Define the range of your table (adjust columns as needed)
    loggingPath = expensesDir & loggingFile

    Set wbLog = Workbooks.Open(loggingPath)
    Set wsLog = wbLog.Sheets("Expense Logging")
    
    With wbLog.Sheets("Expense Logging")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set loggingRange = .Range("A10:M" & lastRow) ' Adjust columns as needed
    End With
    
        ' process each report from the array
    For i = LBound(reportsArray, 1) To UBound(reportsArray, 1)
      If reportsArray(i, 2) = "Reimbursed" Then
        reportID = reportsArray(i, 1)

        ' Find the report ID in the table
        Set foundCell = loggingRange.Find(what:=reportID, LookIn:=xlValues, LookAt:=xlWhole)

        If Not foundCell Is Nothing Then
            ' Update the status to "Reimbursed"
            Application.StatusBar = "Setting Report " & foundCell.Offset(0, 2).Value & "as reimbursed!"
            foundCell.Offset(0, 10).Value = "Reimbursed"
        End If

      ElseIf reportsArray(i, 2) = "Archived" Then
        Application.StatusBar = "Processing report " & i & " of " & numReports & ": Requesting expense information for Report " & reportID & " - Week " & reportsArray(i, 4) & " - " & reportsArray(i, 3) & " from Expensify..."
        Debug.Print "Processing report " & i & " of " & numReports & ": Requesting expense information for Report " & reportID & " - Week " & reportsArray(i, 4) & " - " & reportsArray(i, 3) & " from Expensify..."
                        
        requestType = "csv"
        fileType = "csv"
        reportID = reportsArray(i, 1)
        reportStatus = reportsArray(i, 2)
        submittedDate = reportsArray(i, 5)
        createdDate = reportsArray(i, 6)
        reportTotal = reportsArray(i, 7)
        
        'Get report filename
        reportsFileResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)
        
        'download csv info into an array
        Debug.Print "Processing report " & i & " of " & numReports & ": Grabbing data from Expensify..."
        Application.StatusBar = "Processing report " & i & " of " & numReports & ": Grabbing data from Expensify..."
                
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
        Debug.Print "Processing report " & i & " of " & numReports & ": Producing ESL Expense Report..."
        Application.StatusBar = "Processing report " & i & " of " & numReports & ": Producing ESL Expense Report..."
        Application.ScreenUpdating = False
        Call expensesConversionDataArr(reportsResult, dataArr, wkNr, customer, name, ESLExpenseTemp, expensesDir, saveDirectory, reportName, serialNumber, systemType, reasonForTrip, expectedValue, repCurr, feesTotal, mileagePayout, checkCell, mappedCategory)
        
        'Close the Workbook
        If closeReport = "Yes" Then
            Workbooks(reportName & ".xlsx").Close SaveChanges:=False
        End If
        Application.ScreenUpdating = True
        
        ' Request pdf for report
        Debug.Print "Processing report " & i & " of " & numReports & ": Requesting PDF and getting filename from Expensify..."
        Application.ScreenUpdating = True
        Application.StatusBar = "Processing report " & i & " of " & numReports & ": Requesting PDF and getting filename from Expensify..."
        requestType = "PDF"
        fileType = "pdf"
        filePath = saveDirectory & reportName & "." & fileType
        reportsResult = expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID)

        ' download PDF
        Debug.Print "Processing report " & i & " of " & numReports & ": Downloading PDF..."
        Application.ScreenUpdating = True
        Application.StatusBar = "Processing report " & i & " of " & numReports & ": Downloading PDF..."
        downloadresponse = expensifyAPIdownload(reportsResult, url, userID, userSecret, fileType, requestType, filePath)
        
        ' Save totals to log file
        Application.StatusBar = "Saving expense data to log file..."
                
        loggingPath = expensesDir & loggingFile

        Set wbLog = Workbooks.Open(loggingPath)
        Set wsLog = wbLog.Sheets("Expense Logging")

        logRow = 10
        
        ' Check if reportID already exists and update that row only
        
        ' Find the last row with data (assuming headers are in row 1)
        lastRow = wsLog.Cells(logRow, 1).End(xlDown).Row
        
        If wsLog.Range("A" & logRow).Value <> "" Then ' Check if first cell of log is empty
        
        ' Look for existing reportID in column B (assuming reportID is in column B)
            Set foundRow = wsLog.Range("B" & logRow & ":B" & lastRow).Find(reportID, LookIn:=xlValues, SearchOrder:=xlByRows)
        
            If foundRow Is Nothing Then  ' reportID not found, insert new row
                Debug.Print "No existing ReportID found, inserting new Row"
                wsLog.Rows(logRow).Insert xlShiftDown
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
        wsLog.Range("G" & logRow).Value = checkCell
        If feesTotal <> 0 Then
          wsLog.Range("I" & logRow).Value = feesTotal
        End If
        wsLog.Range("K" & logRow).Value = repCurr
        wsLog.Range("C3").Value = repCurr
        If mileagePayout <> 0 Then
          wsLog.Range("J" & logRow).Value = mileagePayout
        End If
        wsLog.Range("L" & logRow).Value = "Submitted"
        wsLog.Hyperlinks.Add Range("M" & logRow), Address:="file:///" & saveDirectory & reportName & ".xlsx", TextToDisplay:="Link"
        wsLog.Hyperlinks.Add Range("N" & logRow), Address:="file:///" & saveDirectory & reportName & ".pdf", TextToDisplay:="Link"

        ' Calculate difference
        wsLog.Range("H" & logRow).Value = checkCell - expectedValue
        
        ' Calculate allowed range based on error set below
        withinPercent = 5 ' percentage error allowed
      
        lowerLimit = expectedValue * (1 - (withinPercent / 100)) ' 5% less than expected value
        upperLimit = expectedValue * (1 + (withinPercent / 100)) ' 5% more than expected value
      
        If checkCell < lowerLimit Or checkCell > upperLimit Then ' Check if within 5% range
          MsgBox "The ESL Total (" & checkCell & ") is not within 5% of the Expensify Report Total (" & expectedValue & "). Please check before submitting!", vbExclamation
        End If
        
        With wsLog.ListObjects("Table1").Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsLog.Range("Table1[Date Submitted]"), Order:=2
            .Header = xlYes
            .Apply
        End With
        
        Workbooks(loggingFile).RefreshAll

        'Save the logbook
        Workbooks(loggingFile).Save
        
        'Send Email if required
        
        If createEmail = "Yes" Then
        Application.StatusBar = "Creating Email..."
        Call Send_Email(wkNr, customer, reasonForTrip, systemType, saveDirectory, reportName, serialNumber, name, autoSend)
        End If
      End If
    Next i
    
    'Restart Onedrive
    If stopOneDrive = "Yes" Then
      StartOneDriveSync
      startTime = Timer
      Do While Timer < startTime + 5
        Application.StatusBar = "Restarting OneDrive!"
        DoEvents ' Yield processing time to other applications
      Loop
    End If
        
    ' ' Check for any reimbursed reports
    
    ' Application.StatusBar = "Checking for any new reimbursed Reports..."
    ' ' Define the range of your table (adjust columns as needed)
    
    ' requestType = "checkReimbursed"
    ' GoTo Start
    ' loggingPath = expensesDir & loggingFile

    ' Set wbLog = Workbooks.Open(loggingPath)
    ' Set wsLog = wbLog.Sheets("Expense Logging")
    
    ' With wbLog.Sheets("Expense Logging")
    '     lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    '     Set loggingRange = .Range("A10:M" & lastRow) ' Adjust columns as needed
    ' End With

    ' ' Iterate through the report IDs
    ' For n = LBound(reportsArray, 1) To UBound(reportsArray, 1)
    '   If reportsArray(n, 2) = "Reimbursed" Then
    '     reportID = reportsArray(n, 1)

    '     ' Find the report ID in the table
    '     Set foundCell = loggingRange.Find(what:=reportID, LookIn:=xlValues, LookAt:=xlWhole)

    '     If Not foundCell Is Nothing Then
    '         ' Update the status to "Reimbursed"
    '         Application.StatusBar = "Setting Report " & foundCell.Offset(0, 2).Value & "as reimbursed!"
    '         foundCell.Offset(0, 10).Value = "Reimbursed"
    '     End If
    '   End If
  
  'Restart Onedrive
  If stopOneDrive = "Yes" Then
    StartOneDriveSync
    startTime = Timer
    Do While Timer < startTime + 5
      Application.StatusBar = "Restarting OneDrive!"
      DoEvents ' Yield processing time to other applications
    Loop
  End If
  
  'Show completed and wait 5 seconds until clear
  startTime = Timer
  
  Do While Timer < startTime + 5
      Application.StatusBar = "Complete!"
      DoEvents ' Yield processing time to other applications
  Loop

Application.StatusBar = ""


End If

End Sub
