VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSettings 
   Caption         =   "Settings"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7716
   OleObjectBlob   =   "UserFormSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCreateEmail_Change()

End Sub

Private Sub chkMatchTol_Change()

End Sub

Private Sub cmdVisitExpensifyCreds_Click()
Dim webURL As String
    webURL = "https://www.expensify.com/tools/integrations/"

    On Error Resume Next ' Handle errors like invalid address or no internet
    ThisWorkbook.FollowHyperlink Address:=webURL
    If Err.Number <> 0 Then
        MsgBox "Could not open the link:" & vbCrLf & webURL & vbCrLf & _
               "Please check the address and your internet connection.", vbExclamation, "Hyperlink Error"
        Err.Clear
    End If
    On Error GoTo 0 ' Turn error handling back to normal
End Sub


Private Sub MultiPage1_Change()

End Sub

' --- UserForm Module: UserFormSettings ---


Private Sub UserForm_Initialize()
    ' Call the function in the module to load settings into THIS form instance (Me)
        Dim itemArray As Variant

    ' Create an array of items
    itemArray = Array("No", "Individual")

    ' Assign the array to the ComboBox's List property
    Me.cboCreateEmail.List = itemArray
    
    modSettings.LoadAllSettingsIntoForm Me
End Sub

Private Sub cmdBrowseExpensesDir_Click()
    Dim selectedPath As String
    ' Call the helper in the module
    selectedPath = modSettings.BrowseForFolder("Select Expenses Directory", Me.txtExpensesDir.Text) & "\"
    If selectedPath <> "" Then Me.txtExpensesDir.Text = selectedPath
End Sub

Private Sub cmdBrowseESLTemplate_Click()
    Dim selectedPath As String
    ' Call the helper in the module - specify filter for Excel files
    selectedPath = modSettings.BrowseForFile("Select ESL Template File (Full Path)", "Excel Files (*.xls*),*.xls*", Me.txtESLTemplate.Text)
    If selectedPath <> "" Then Me.txtESLTemplate.Text = selectedPath
End Sub

Private Sub cmdBrowseLoggingFile_Click()
    Dim selectedPath As String
    ' Call the helper in the module - specify filter for log/text files
    selectedPath = modSettings.BrowseForFile("Select Logging File (Full Path)", "Log/Text Files (*.log;*.txt),*.log;*.txt", Me.txtLoggingFile.Text)
    If selectedPath <> "" Then Me.txtLoggingFile.Text = selectedPath
End Sub

Private Sub chkUseFees_Click()
    ' This directly affects UI state, so keep it here
    UpdateFeeControls ' Call helper sub below
End Sub

Public Sub UpdateFeeControls()
    ' Helper sub within the form for UI updates. Made Public so LoadAllSettings can call it.
    On Error Resume Next ' Avoid errors if controls don't exist during design time maybe
    Me.txtDefaultFees.Visible = Me.chkUseFees.Value
    Me.labelFees.Visible = Me.chkUseFees.Value
    Me.txtDefaultFees.Enabled = Me.chkUseFees.Value
    On Error GoTo 0
End Sub

Private Sub cmdOK_Click()
Dim startTime As Double
    ' Call the function in the module to save settings FROM THIS form instance (Me)
    If modSettings.SaveAllSettingsFromForm(Me) Then
        modSettings.LoadSettingsIntoMemory
        Unload Me ' Close form only if save was successful
        Debug.Print "Settings Saved."
        Application.StatusBar = "Settings Saved!"
        startTime = Timer
        Do While Timer < startTime + 5
          DoEvents ' Yield processing time to other applications
        Loop
        
    Else
        ' Optional: Keep form open if save failed due to validation or error
    End If
  Application.StatusBar = ""
End Sub

Private Sub cmdCancel_Click()
    modSettings.LoadSettingsIntoMemory
    Unload Me ' Close without saving changes made on the form
End Sub

' --- Optional: Add ToolTips in Initialize for better UX ---
Private Sub SetToolTips()
    On Error Resume Next ' Ignore errors if controls don't exist
    Me.txtExpensesDir.ControlTipText = "Folder where expense-related files will be processed or saved."
    Me.cmdBrowseExpensesDir.ControlTipText = "Browse for the expenses directory."
    Me.txtESLTemplate.ControlTipText = "Full path to the ESL Template Excel file."
    Me.cmdBrowseESLTemplate.ControlTipText = "Browse for the ESL Template file."
    Me.txtLoggingFile.ControlTipText = "Full path to the file used for logging process details."
    Me.cmdBrowseLoggingFile.ControlTipText = "Browse for the logging file."
    Me.txtUserID.ControlTipText = "Your Expensify Integrations User ID."
    Me.txtUserSecret.ControlTipText = "Your Expensify Integrations User Secret."
    Me.txtPolicyID.ControlTipText = "The Policy ID for your Expensify Workspace."
    Me.chkUseFees.ControlTipText = "Check this box to automatically calculate fees based on a percentage."
    Me.txtDefaultFees.ControlTipText = "Enter the default percentage (e.g., 2.5) for fee calculation."
    Me.cmdOK.ControlTipText = "Save the current settings and close this window."
    Me.cmdCancel.ControlTipText = "Close this window without saving any changes."
    On Error GoTo 0
End Sub

' Call SetToolTips from UserForm_Initialize if desired
' Private Sub UserForm_Initialize()
'     modSettings.LoadAllSettingsIntoForm Me
'     SetToolTips ' Call the tooltip setup
' End Sub
Private Sub UserForm_Terminate()
    modSettings.LoadSettingsIntoMemory
    Unload Me ' Close without saving changes made on the form
End Sub
