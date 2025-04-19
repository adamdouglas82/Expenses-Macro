Attribute VB_Name = "URLEncodeModule"
Function URLEncode(str As String) As String
  URLEncode = WorksheetFunction.EncodeURL(str)
End Function
