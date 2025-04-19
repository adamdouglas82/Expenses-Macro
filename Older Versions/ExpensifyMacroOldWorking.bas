Attribute VB_Name = "ExpensifyMacroOldWorking"
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


Function URLEncode(str As String) As String
  URLEncode = WorksheetFunction.EncodeURL(str)
End Function

Function StopOneDriveSync()

Shell """" & Environ$("ProgramFiles") & "\Microsoft OneDrive\onedrive.exe" & """" & " /shutdown"

End Function

Function StartOneDriveSync()

Shell """" & Environ$("ProgramFiles") & "\Microsoft OneDrive\onedrive.exe" & """" & " /background"

End Function

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
Dim numReports As Integer
Dim reportStatus As String
Dim mappedCategory As String
Dim upgrade As String
Dim update As String
Dim useFees As String
Dim defaultFees As String
Dim foundRow As Range
Dim stopOneDrive As String

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
    If reDownload = "All" Or reDownload = "Submitted" Then
        ActiveSheet.Range("M2").Value = "Submitted"
    End If
    requestType = "csv"
    reportsResult = expensifyAPIdownload(reportsResult, url, userID, userSecret, fileType, requestType, filePath)
    
    ' determine if empty then get reports list
    
    emptyIdentifier = "$empty$"
    reportsResult = emptyIdentifier & Trim(reportsResult)
    
    ' remove trailing comma or empty character from string
    reportsResult = Left(reportsResult, Len(reportsResult) - 1)
        
    If Left(reportsResult, Len(emptyIdentifier) + 1) = emptyIdentifier Then
        MsgBox ("No closed reports found!" & vbNewLine & vbNewLine & "Please check for closed reports in Expensify." & vbNewLine & vbNewLine & "Checking for reimbursed reports...")
        requestType = "checkReimbursed"
        GoTo Start
    End If
    
    ' restore normal result string
    reportsResult = Mid(reportsResult, Len(emptyIdentifier) + 1)
    
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
    
        ' process each report from the array
    For i = LBound(reportsArray, 1) To UBound(reportsArray, 1)
      If reportsArray(i, 2) = "Archived" Or reportsArray(i, 2) = "Reimbursed" Then
        Application.StatusBar = "Processing report " & i & " of " & numReports & ": Requesting expense information for Report " & reportID & " - Week " & reportsArray(i, 4) & " - " & reportsArray(i, 3) & " from Expensify..."
        Debug.Print "Processing report " & i & " of " & numReports & ": Requesting expense information for Report " & reportID & " - Week " & reportsArray(i, 4) & " - " & reportsArray(i, 3) & " from Expensify..."
        requestType = "csv"
        fileType = "csv"
        reportID = reportsArray(i, 1)
        submittedDate = reportsArray(i, 5)
        reportStatus = reportsArray(i, 2)
        
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
        
    ' Check for any reimbursed reports
    
    Application.StatusBar = "Checking for any new reimbursed Reports..."
    ' Define the range of your table (adjust columns as needed)
    
    requestType = "checkReimbursed"
    GoTo Start
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
  Dim templatePath As String
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
   
  ' get template path
  templatePath = expensesDir & ESLExpenseTemp

  'find last backslash in templatePath
  lastBackslash = InStrRev(templatePath, "\")

  ' Extract the directory path by taking everything up to the last backslash
  directoryPath = Left(templatePath, lastBackslash)
  
  ' Create new workbook from template
  Set wbNew = Workbooks.Add(template:=templatePath)
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

  expenseType = dataArr(dataRow, catCol)
  If expenseType = "Mileage" Then ' Process only Mileage Entries
    'mileageTotal = mileageTotal + CDbl(dataArr(dataRow, mileCol)) ' Calculate mileage total (assuming numeric value)
    mileageRate = dataArr(dataRow, milerateCol) / 100 ' Find mileage rate and divide by 100 due to expensify report bug

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
    
    ' Insert Miles
    wsMileage.Range("H" & mileageRow).Value = dataArr(dataRow, mileCol)
    
    ' Copy over Mileage rate if not GBP or USD
    If dataArr(2, repCurrCol) <> "GBP" And dataArr(2, repCurrCol) <> "USD" Then
      wsOutput.Range("P52").Value = mileageRate ' Output mileage rate if currency is not GBP or USD
    End If
    
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
    expenseType = dataArr(dataRow, catCol) ' column containing the expense type
    For col = 6 To 14 ' Loop through category name cells (F16:N16)
      If expenseType = wsOutput.Range(Cells(16, col).Address) Then ' Check for matching expense type
        wsOutput.Range(Cells(outputRow, col).Address) = dataArr(dataRow, expamtCol) ' Write expense amount to category column
        Exit For ' Exit loop after finding a match
      End If
    Next col
    
    ' Check if currency matches report currency
      If dataArr(dataRow, expCurrCol) <> dataArr(2, repCurrCol) And IsNumeric(dataArr(dataRow, feesCol)) And dataArr(dataRow, feesCol) <> 0 Then
        wsOutput.Range("S" & outputRow).Formula = "=ROUND(O" & outputRow & "*(" & dataArr(dataRow, feesCol) & "/100)*Q" & outputRow & ",2)"
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

End Function
Function expensifyAPIRequest(requestType, reportID, userID, userSecret, url, reportName, policyID, Optional reDownload, Optional useFees, Optional defaultFees)

Dim postdata As String
Dim httpRequest As Object
Dim templateContent As String
Dim templateContentURL As String
Dim fullSend As String
Dim expensifyRequestReturn As String
Dim fileType As String
Dim reportFilters As String
Dim reportState As String


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
    
    ' set redownload expenses or not
    If reDownload = "All" Then
        reportFilters = """startDate"":""2024-01-10""}},"
        reportState = """reportState"":""ARCHIVED,REIMBURSED"","
        
    ElseIf reDownload = "Submitted" Then
        reportFilters = """startDate"":""2024-01-10""}},"
        reportState = """reportState"":""ARCHIVED"","
    Else
        reportFilters = """startDate"":""2024-01-10"",""markedAsExported"":""Expensify Export""}},"
        reportState = """reportState"":""ARCHIVED,REIMBURSED"","
    End If
    
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

    templateContent = "<#compress><#list reports as report>${report.reportID};${report.status};${report.customField.Customer};${report.customField.Week};${report.submitted}" & vbCrLf & "</#list></#compress>"

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
    
    If useFees = "Yes" Then
        useFees = ",{""name"": ""Fees"",""type"":""text"",""setRequired"": false,""defaultValue"":""" & defaultFees & """}" & "]}}"
    Else
        useFees = "]}}"
    End If
    
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
                                            & useFees


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



