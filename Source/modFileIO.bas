Attribute VB_Name = "modFileIO"
Option Explicit

'=========================================================
'
' VB6FastIni
'
' Fast file loading routines.
'
'=========================================================

Public Function ReadTextFile(ByVal FileName As String) As String

    Dim FileNumber As Integer

    On Error GoTo ErrorHandler

    If Dir$(FileName, vbNormal) = "" Then Exit Function

    FileNumber = FreeFile

    Open FileName For Binary Access Read As #FileNumber

    If LOF(FileNumber) > 0 Then
        ReadTextFile = Space$(LOF(FileNumber))
        Get #FileNumber, , ReadTextFile
    End If

    Close #FileNumber

    Exit Function

ErrorHandler:

    If FileNumber <> 0 Then
        Close #FileNumber
    End If

    Err.Raise Err.Number, _
              "ReadTextFile", _
              Err.Description

End Function


Public Sub WriteTextFile(ByVal FileName As String, _
                         ByVal Text As String)

    Dim FileNumber As Integer

    On Error GoTo ErrorHandler

    FileNumber = FreeFile

    Open FileName For Binary Access Write As #FileNumber

    Put #FileNumber, , Text

    Close #FileNumber

    Exit Sub

ErrorHandler:

    If FileNumber <> 0 Then
        Close #FileNumber
    End If

    Err.Raise Err.Number, _
              "WriteTextFile", _
              Err.Description

End Sub
