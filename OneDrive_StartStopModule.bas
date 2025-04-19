Attribute VB_Name = "OneDrive_StartStopModule"
Function StopOneDriveSync()

Shell """" & Environ$("ProgramFiles") & "\Microsoft OneDrive\onedrive.exe" & """" & " /shutdown"

End Function

Function StartOneDriveSync()

Shell """" & Environ$("ProgramFiles") & "\Microsoft OneDrive\onedrive.exe" & """" & " /background"

End Function
