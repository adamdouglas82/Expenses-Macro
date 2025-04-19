Attribute VB_Name = "modSettings"
Option Explicit

' --- Standard Module: modSettings ---

' --- Public Constants (Accessible Project-Wide) ---
Public Const SETTINGS_SHEET_NAME As String = "AppSettings" ' Hidden sheet name
' Cell addresses for setting VALUES (Column A has labels)
Public Const EXPENSE_PATH_CELL As String = "B1"
Public Const ESL_TEMPLATE_PATH_CELL As String = "B2" ' Store FULL path now
Public Const LOGGING_FILE_PATH_CELL As String = "B3" ' Store FULL path now
Public Const USER_ID_CELL As String = "B4"
Public Const USER_SECRET_CELL As String = "B5"
Public Const POLICY_ID_CELL As String = "B6"
Public Const USE_FEES_CELL As String = "B7" ' Store "Yes" or "No"
Public Const DEFAULT_FEES_CELL As String = "B8"
Public Const POLICY_CONFIGURED_CELL As String = "B9"
Public Const AUTO_RUN_CELL As String = "B10"
Public Const CLOSE_REPORT_CELL As String = "B11"
Public Const CREATE_EMAIL_CELL As String = "B12"
Public Const AUTO_SEND_CELL As String = "B13"
Public Const STOP_ONEDRIVE_CELL As String = "B14"
Public Const TOLERANCE_CELL As String = "B15"
Public Const LOGGING_FILENAME_ONLY_CELL As String = "B16"

Public expensesDir As String
Public ESLExpenseTemp As String
Public loggingFile As String
Public loggingFilenameOnly As String
Public userID As String
Public userSecret As String ' Consider security implications if storing/loading this
Public policyID As String
Public useFees As Boolean
Public defaultFees As Double ' Use Double for percentage if appropriate
Public policyConfigured As Boolean
Public AutoRun As Boolean
Public closeReport As Boolean
Public createEmail As String
Public autoSend As Boolean
Public stopOneDrive As Boolean
Public tolerance As Double


Public Sub LoadSettingsIntoMemory()
    Dim ws As Worksheet
    Set ws = GetSettingsSheet(ThisWorkbook) ' Use your existing helper

    If ws Is Nothing Then
        MsgBox "Cannot load settings into memory: Settings sheet '" & SETTINGS_SHEET_NAME & "' not found.", vbCritical
        ' Clear variables to ensure they don't hold old values
        expensesDir = ""
        ESLExpenseTemp = ""
        loggingFile = ""
        userID = ""
        userSecret = ""
        policyID = ""
        useFees = False
        defaultFees = 0
        policyConfigured = False
        loggingFilenameOnly = ""
        Exit Sub
    End If

    On Error Resume Next ' Handle errors reading individual cells

    ' Read from sheet cells and assign to public variables
    expensesDir = CStr(ws.Range(EXPENSE_PATH_CELL).Value)
    ESLExpenseTemp = CStr(ws.Range(ESL_TEMPLATE_PATH_CELL).Value)
    loggingFile = CStr(ws.Range(LOGGING_FILE_PATH_CELL).Value)
    userID = CStr(ws.Range(USER_ID_CELL).Value)
    userSecret = CStr(ws.Range(USER_SECRET_CELL).Value)
    policyID = CStr(ws.Range(POLICY_ID_CELL).Value)
    policyConfigured = CBool(ws.Range(POLICY_CONFIGURED_CELL).Value)
    useFees = CBool(ws.Range(USE_FEES_CELL).Value)
    If useFees Then
        ' Val converts text to number, returns 0 if not convertible
        defaultFees = Val(CStr(ws.Range(DEFAULT_FEES_CELL).Value))
    End If
    AutoRun = CBool(ws.Range(AUTO_RUN_CELL).Value)
    closeReport = CBool(ws.Range(CLOSE_REPORT_CELL).Value)
    createEmail = CStr(ws.Range(CREATE_EMAIL_CELL).Value)
    autoSend = CBool(ws.Range(AUTO_SEND_CELL).Value)
    stopOneDrive = CBool(ws.Range(STOP_ONEDRIVE_CELL).Value)
    tolerance = CStr(ws.Range(TOLERANCE_CELL).Value)
    loggingFilenameOnly = CStr(ws.Range(LOGGING_FILENAME_ONLY_CELL).Value)
     
     If Err.Number <> 0 Then
         MsgBox "An error occurred reading settings from '" & SETTINGS_SHEET_NAME & "'." & vbCrLf & _
                "Some settings may not be loaded correctly.", vbExclamation
         Err.Clear
     End If
     On Error GoTo 0
     Set ws = Nothing
     Debug.Print "Settings loaded into memory variables." ' For testing
End Sub

Public Sub OpenSettingsForm()
    ' Purpose: Shows the main settings UserForm.
    On Error Resume Next ' Basic error handling
    UserFormSettings.Show

    ' Optional: Check if showing the form failed (e.g., form name incorrect)
    If Err.Number <> 0 Then
        MsgBox "Error opening settings form: " & Err.Description, vbCritical, "Error"
        Err.Clear
    End If
    On Error GoTo 0 ' Turn error handling off

End Sub

' --- Core Settings Sheet Handling ---
Public Function GetSettingsSheet(wb As Workbook) As Worksheet
    ' Ensures the settings sheet exists and is very hidden
    Dim ws As Worksheet
    On Error Resume Next ' Handle error if sheet doesn't exist
    Set ws = wb.Sheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0 ' Turn error handling back on

    If ws Is Nothing Then ' Sheet doesn't exist, create it
        On Error Resume Next ' Handle potential errors during sheet creation/naming
        Application.ScreenUpdating = False ' Prevent screen flicker
        ' Important: Adding sheets might fail if Workbook Structure is protected
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        If Err.Number <> 0 Then GoTo CreateFail ' Exit if sheet add failed

        ws.name = SETTINGS_SHEET_NAME
        If Err.Number <> 0 Then GoTo CreateFail ' Exit if rename failed

        ws.Visible = xlSheetVeryHidden ' Make it super hidden

        ' Optional: Add Labels in Column A for clarity on the sheet
        ws.Range("A1:A16").Value = Application.Transpose(Array( _
            "Expenses Dir:", "ESL Template Path:", "Logging File Path:", _
            "User ID:", "User Secret:", "Policy ID:", "Use Fees:", "Default Fees:", _
            "Policy Configured:", "AutoRun:", "Close Report:", "Create Email:", "Auto Send Email:", _
            "Stop One Drive:", "Match Tolerance:", "Logging Filename Only:"))
        ws.columns("A").AutoFit
        
        ws.Range("B1").Value = "" ' Expenses Dir
        ws.Range("B2").Value = "" ' ESL Template Path (Suggest blank or specific default)
        ws.Range("B3").Value = "" ' Logging File Path
        ws.Range("B4").Value = "" ' User ID
        ws.Range("B5").Value = "" ' User Secret (Always leave blank)
        ws.Range("B6").Value = "" ' Policy ID
        ws.Range("B7").Value = False ' Use Fees (Boolean False)
        ws.Range("B8").Value = 2  ' Default Fees (Number, e.g., 2.5)
        ws.Range("B9").Value = False ' Policy Configured (Boolean False)
        ws.Range("B10").Value = False ' AutoRun (Default to True?)
        ws.Range("B11").Value = False ' Close Report (Default to True?)
        ws.Range("B12").Value = "NO" ' Create Email (Default to True?)
        ws.Range("B13").Value = False ' Auto Send Email (Default to False?)
        ws.Range("B14").Value = False ' Stop One Drive (Default to False?)
        ws.Range("B15").Value = 2 ' Match Tolerance (Default to 0?)
        ws.Range(LOGGING_FILENAME_ONLY_CELL).Value = ""
        ws.columns("B").AutoFit

CreateFail:
        Application.ScreenUpdating = True
        If ws Is Nothing Then MsgBox "Failed to create settings sheet '" & SETTINGS_SHEET_NAME & "'. Check workbook protection or sheet name validity.", vbCritical
        On Error GoTo 0
    Else ' Sheet exists, ensure it's very hidden
        If ws.Visible <> xlSheetVeryHidden Then
             On Error Resume Next ' Handle if visibility change fails (protection)
             ws.Visible = xlSheetVeryHidden
             On Error GoTo 0
        End If
    End If
    Set GetSettingsSheet = ws ' Return the sheet object (or Nothing if creation failed)
End Function

' --- Check Configuration Status ---
Public Function AreSettingsConfigured() As Boolean
    ' Checks if essential settings seem to be filled in
    Dim wsSettings As Worksheet
    Dim configComplete As Boolean
    configComplete = False ' Assume not configured initially
    Set wsSettings = GetSettingsSheet(ThisWorkbook) ' Use helper

    If Not wsSettings Is Nothing Then
        ' Check key required fields - adjust these as needed for *your* requirements
        If wsSettings.Range(EXPENSE_PATH_CELL).Value <> "" And _
           wsSettings.Range(ESL_TEMPLATE_PATH_CELL).Value <> "" And _
           wsSettings.Range(USER_ID_CELL).Value <> "" And _
           wsSettings.Range(POLICY_ID_CELL).Value <> "" Then
             configComplete = True
        End If
    End If
    AreSettingsConfigured = configComplete
    Set wsSettings = Nothing ' Clean up
End Function

' --- Prompt User for Initial Setup (Called from Workbook_Open or Main Macro) ---
Public Sub PromptForSettings()
    ' Shows message and the settings form modally
    ' Make sure this sub is Public
    MsgBox "Welcome! Initial setup is required. Please configure the application settings.", vbInformation + vbOKOnly, "Setup Required"
    On Error Resume Next ' Handle if form is already open? Unlikely if called modally
    UserFormSettings.Show vbModal ' Show the form
    If Err.Number <> 0 Then MsgBox "Error showing settings form.", vbCritical
    On Error GoTo 0
    'Optional: Re-check after form closes and warn if still not configured
     If Not AreSettingsConfigured() Then
        MsgBox "Warning: Settings were not fully configured.", vbExclamation
        UserFormSettings.Show vbModal ' Show the form
        If Err.Number <> 0 Then MsgBox "Error showing settings form.", vbCritical
        On Error GoTo 0
     End If
End Sub

' --- Load Settings FROM Sheet INTO Form Controls ---
Public Sub LoadAllSettingsIntoForm(frm As UserFormSettings) ' Pass form object
    Dim wsSettings As Worksheet
    Set wsSettings = GetSettingsSheet(ThisWorkbook)

    If Not wsSettings Is Nothing Then
        frm.txtExpensesDir.Text = CStr(wsSettings.Range(EXPENSE_PATH_CELL).Value)
        frm.txtESLTemplate.Text = CStr(wsSettings.Range(ESL_TEMPLATE_PATH_CELL).Value)
        frm.txtLoggingFile.Text = CStr(wsSettings.Range(LOGGING_FILE_PATH_CELL).Value)
        frm.txtUserID.Text = CStr(wsSettings.Range(USER_ID_CELL).Value)
        frm.txtUserSecret.Text = CStr(wsSettings.Range(USER_SECRET_CELL).Value)
        frm.txtPolicyID.Text = CStr(wsSettings.Range(POLICY_ID_CELL).Value)
        frm.chkUseFees.Value = wsSettings.Range(USE_FEES_CELL).Value
        frm.txtDefaultFees.Text = Format(wsSettings.Range(DEFAULT_FEES_CELL).Value, "0.0")
        frm.chkAutoRun.Value = wsSettings.Range(AUTO_RUN_CELL).Value
        frm.chkCloseReport.Value = wsSettings.Range(CLOSE_REPORT_CELL).Value
        frm.cboCreateEmail.Value = wsSettings.Range(CREATE_EMAIL_CELL).Value
        frm.chkAutoSendEmail.Value = wsSettings.Range(AUTO_SEND_CELL).Value
        frm.chkStopOneDrive.Value = wsSettings.Range(STOP_ONEDRIVE_CELL).Value
        frm.chkMatchTol.Value = Format(wsSettings.Range(TOLERANCE_CELL).Value, "0.0")

        ' Ensure dependent UI elements are updated
        frm.UpdateFeeControls ' Call public sub within the form to update UI
    Else
        MsgBox "Could not load settings from sheet '" & SETTINGS_SHEET_NAME & "'.", vbCritical
        ' Optional: Disable form controls if settings cannot be loaded
        ' frm.Enabled = False ' Or disable a specific frame/controls
    End If
    Set wsSettings = Nothing
End Sub

' --- Save Settings FROM Form Controls TO Sheet ---
Public Function SaveAllSettingsFromForm(frm As UserFormSettings) As Boolean ' Pass form object, return success/fail
    Dim wsSettings As Worksheet
    Dim success As Boolean
    Dim filenameOnly As String
    success = False ' Assume failure initially
    Set wsSettings = GetSettingsSheet(ThisWorkbook)

    If Not wsSettings Is Nothing Then
        ' --- Validation ---
        If frm.txtExpensesDir.Text = "" Or frm.txtESLTemplate.Text = "" Or frm.txtLoggingFile.Text = "" Or frm.txtUserID.Text = "" Or frm.txtUserSecret.Text = "" Or frm.txtPolicyID.Text = "" Then
             MsgBox "Please ensure all required fields (paths, user ID, secret, policy ID) are filled.", vbExclamation, "Validation Error"
             GoTo Cleanup ' Exit function, returning False
        End If
        If frm.chkUseFees.Value And (Not IsNumeric(frm.txtDefaultFees.Text) Or frm.txtDefaultFees.Text = "") Then
             MsgBox "Please enter a numeric value for the default fees percentage.", vbExclamation, "Validation Error"
             frm.txtDefaultFees.SetFocus
             GoTo Cleanup ' Exit function, returning False
        End If
        ' Add more specific validation (e.g., check if paths exist?) if desired
        ' --- End Validation ---

        On Error Resume Next ' Handle potential write errors (e.g., sheet protected)
        wsSettings.Range(EXPENSE_PATH_CELL).Value = frm.txtExpensesDir.Text
        wsSettings.Range(ESL_TEMPLATE_PATH_CELL).Value = frm.txtESLTemplate.Text ' Saving full path
        wsSettings.Range(LOGGING_FILE_PATH_CELL).Value = frm.txtLoggingFile.Text ' Saving full path
        wsSettings.Range(USER_ID_CELL).Value = frm.txtUserID.Text
        wsSettings.Range(USER_SECRET_CELL).Value = frm.txtUserSecret.Text ' Save password field value - consider security implications
        wsSettings.Range(POLICY_ID_CELL).Value = frm.txtPolicyID.Text
        wsSettings.Range(USE_FEES_CELL).Value = frm.chkUseFees.Value
        If frm.chkUseFees.Value Then wsSettings.Range(DEFAULT_FEES_CELL).Value = frm.txtDefaultFees.Text
        wsSettings.Range(AUTO_RUN_CELL).Value = frm.chkAutoRun.Value
        wsSettings.Range(CLOSE_REPORT_CELL).Value = frm.chkCloseReport.Value
        wsSettings.Range(CREATE_EMAIL_CELL).Value = frm.cboCreateEmail.Text
        wsSettings.Range(AUTO_SEND_CELL).Value = frm.chkAutoSendEmail.Value
        wsSettings.Range(STOP_ONEDRIVE_CELL).Value = frm.chkStopOneDrive.Value
        wsSettings.Range(TOLERANCE_CELL).Value = frm.chkMatchTol.Text
        
        If InStrRev(frm.txtLoggingFile.Text, "\") > 0 Then
          filenameOnly = Mid$(frm.txtLoggingFile.Text, InStrRev(frm.txtLoggingFile.Text, "\") + 1)
        ElseIf frm.txtLoggingFile.Text <> "" Then ' Handle if only filename was entered?
          filenameOnly = frm.txtLoggingFile.Text ' Use input as filename if no path separator
        Else
          filenameOnly = "" ' Set to blank if full path is blank
        End If
        wsSettings.Range(LOGGING_FILENAME_ONLY_CELL).Value = filenameOnly ' *** SAVE Filename to B16 ***

        If Err.Number <> 0 Then
            MsgBox "An error occurred while saving settings to the '" & SETTINGS_SHEET_NAME & "' sheet." & vbCrLf & _
                   "Please ensure the sheet is not protected.", vbExclamation, "Save Error"
            Err.Clear
             ' Success remains False
        Else
            success = True ' Save was successful
        End If
        On Error GoTo 0
    Else
        MsgBox "Could not access settings sheet '" & SETTINGS_SHEET_NAME & "'. Settings not saved.", vbCritical
         ' Success remains False
    End If

Cleanup:
    SaveAllSettingsFromForm = success ' Return True if saved, False otherwise
    Set wsSettings = Nothing
End Function


Public Function SaveSpecificSetting(settingCellAddress As String, newValue As Variant) As Boolean
    ' Purpose: Saves a single value back to the settings sheet.
    ' Used for updating flags or individual settings programmatically, like the Policy Configured flag.
    ' Input:  settingCellAddress - The cell address (e.g., "B9") from constants.
    '         newValue - The value to save (e.g., True).
    ' Output: Returns True on success, False on failure.

    Dim ws As Worksheet
    Dim success As Boolean
    success = False ' Assume failure initially

    ' Get the settings sheet using the helper function
    Set ws = GetSettingsSheet(ThisWorkbook)

    If Not ws Is Nothing Then
        ' Sheet exists, proceed to write
        On Error Resume Next ' Handle potential write errors (e.g., sheet protected)

        ' Write the new value to the specified cell address on that sheet
        ws.Range(settingCellAddress).Value = newValue

        ' Check if the write operation caused an error
        If Err.Number = 0 Then
            success = True ' Write succeeded
            Debug.Print "Saved value '" & CStr(newValue) & "' to cell " & settingCellAddress & " on sheet " & ws.name
        Else
            ' Write failed, report error
            MsgBox "Error saving setting to cell " & settingCellAddress & " on sheet '" & SETTINGS_SHEET_NAME & "'." & vbCrLf & _
                   "Sheet might be protected. Error: " & Err.Description, vbExclamation, "Save Setting Error"
            Err.Clear ' Clear the error
        End If
        On Error GoTo 0 ' Turn default error handling back on
    Else
        ' Critical error: Settings sheet itself couldn't be found or created.
        MsgBox "Cannot save setting: Settings sheet '" & SETTINGS_SHEET_NAME & "' not found or could not be created.", vbCritical
        ' success remains False
    End If

    SaveSpecificSetting = success ' Return True if save was successful, False otherwise
    Set ws = Nothing ' Clean up worksheet object
End Function

' --- General Helper Functions for Browse ---
Public Function BrowseForFolder(Optional title As String = "Select Folder", Optional initialPath As String = "") As String
    Dim fldrPicker As FileDialog
    Dim selectedPath As String
    selectedPath = "" ' Default to empty string (cancel)

    Set fldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With fldrPicker
        .title = title
        .AllowMultiSelect = False
        On Error Resume Next ' Handle error if initialPath is invalid
        If initialPath <> "" And Dir(initialPath, vbDirectory) <> "" Then
            .InitialFileName = initialPath
        End If
        On Error GoTo 0 ' Back to normal error handling

        If .Show = -1 Then ' -1 means user clicked OK
             selectedPath = .SelectedItems(1)
        End If
    End With
    BrowseForFolder = selectedPath
    Set fldrPicker = Nothing ' Clean up
End Function

Public Function BrowseForFile(Optional title As String = "Select File", Optional filter As String = "All Files (*.*),*.*", Optional initialPath As String = "") As String
     Dim filePicker As FileDialog
     Dim selectedPath As String
     selectedPath = "" ' Default to empty string (cancel)

     Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
     With filePicker
         .title = title
         .AllowMultiSelect = False
         ' Set up filters
         If filter <> "" Then
             On Error Resume Next ' Handle potential errors with filter format
             .Filters.Clear
             .Filters.Add Split(filter, ",")(0), Split(filter, ",")(1)
             .Filters.Add "All Files", "*.*" ' Always add All Files filter
             If Err.Number <> 0 Then ' If filter format was bad, just use All Files
                 .Filters.Clear
                 .Filters.Add "All Files", "*.*"
                 Err.Clear
             End If
             On Error GoTo 0
         Else
             .Filters.Clear
             .Filters.Add "All Files", "*.*"
         End If

        ' Set initial view location
         On Error Resume Next ' Handle error if initialPath is invalid
         If initialPath <> "" Then
              ' Check if it's a directory or includes a filename
              If Dir(initialPath, vbDirectory) <> "" Then
                  .InitialFileName = initialPath & "\" ' Start in this folder
              ElseIf Dir(initialPath) <> "" Then
                  .InitialFileName = initialPath ' Suggest this file
              End If
         End If
         On Error GoTo 0 ' Back to normal error handling

         If .Show = -1 Then ' -1 means user clicked OK
             selectedPath = .SelectedItems(1)
         End If
     End With
     BrowseForFile = selectedPath
     Set filePicker = Nothing ' Clean up
 End Function

