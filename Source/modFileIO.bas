Attribute VB_Name = "modFileIO"
Option Explicit
Public Function ReadTextFile(ByVal FileName As String) As String
Dim FileNumber As Integer
Dim FileIsOpen As Boolean
Dim ErrorNumber As Long
Dim ErrorDescription As String

On Error GoTo ErrorHandler

If Dir$(FileName, vbNormal) = "" Then
    Exit Function
End If

FileNumber = FreeFile

Open FileName For Binary Access Read As #FileNumber
FileIsOpen = True

If LOF(FileNumber) > 0 Then
    ReadTextFile = Space$(LOF(FileNumber))
    Get #FileNumber, , ReadTextFile
End If

Close #FileNumber
FileIsOpen = False

Exit Function

ErrorHandler:

ErrorNumber = Err.Number
ErrorDescription = Err.Description
On Error Resume Next
If FileIsOpen Then Close #FileNumber
On Error GoTo 0

Err.Raise ErrorNumber, "ReadTextFile", ErrorDescription

End Function

Public Sub WriteTextFile(ByVal FileName As String, ByVal Text As String)
Dim FileNumber As Integer
Dim FileIsOpen As Boolean
Dim ErrorNumber As Long
Dim ErrorDescription As String

On Error GoTo ErrorHandler

FileNumber = FreeFile

Open FileName For Output As #FileNumber
FileIsOpen = True

Print #FileNumber, Text;

Close #FileNumber
FileIsOpen = False

Exit Sub

ErrorHandler:

ErrorNumber = Err.Number
ErrorDescription = Err.Description
On Error Resume Next
If FileIsOpen Then Close #FileNumber
On Error GoTo 0

Err.Raise ErrorNumber, "WriteTextFile", ErrorDescription

End Sub
