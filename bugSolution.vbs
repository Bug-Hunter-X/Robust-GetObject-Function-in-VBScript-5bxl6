Function GetObject(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = CreateObject(progID)
    If Err.Number <> 0 Then 'Handle CreateObject errors
        MsgBox "Error creating object: " & Err.Description, vbCritical
        Set obj = Nothing 'Important: Set the object to Nothing to prevent memory leaks
    End If
  End If
  Set GetObject = obj
End Function

'Example Usage:
On Error GoTo ErrorHandler
Set myExcel = GetObject("Excel.Application")

If Not myExcel Is Nothing Then
  MsgBox "Excel Application found."
  myExcel.Visible = True
  myExcel.Quit
  Set myExcel = Nothing
Else
  MsgBox "Excel Application not found or failed to create."
End If

Exit Sub

ErrorHandler:
MsgBox "An error occurred: " & Err.Description, vbCritical
Err.Clear
