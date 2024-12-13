Option Explicit

' Function to demonstrate safe type checking
Function AddNumbers(num1, num2)
  If IsNumeric(num1) And IsNumeric(num2) Then
    AddNumbers = CDbl(num1) + CDbl(num2) 'Explicit type conversion
  Else
    Err.Raise 13, , "Invalid input: Arguments must be numbers." 'Error handling
  End If
End Function

' Example of early binding for better error handling
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check if file exists before trying to access it
If objFSO.FileExists("myFile.txt") Then
    'Proceed with file operations
    WScript.Echo "File exists"
Else
    WScript.Echo "File does not exist"
End If

Set objFSO = Nothing

'Example of explicit type conversion and error handling
Dim strValue As String
Dim intValue As Integer

strValue = "123"

On Error Resume Next
intValue = CInt(strValue)
If Err.Number <> 0 Then
  WScript.Echo "Error converting string to integer: " & Err.Description
  Err.Clear
End If
On Error GoTo 0

WScript.Echo "Addition: " & AddNumbers(10, 20)
WScript.Echo "Addition: " & AddNumbers("10", "20")
WScript.Echo "Addition: " & AddNumbers("abc", 20) 