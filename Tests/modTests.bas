Attribute VB_Name = "modTests"
Option Explicit

Private m_TempFolder As String
Private m_RunId As String
Private m_Assertions As Long
Private m_Failures As Long

Public Sub Main()

m_TempFolder = Environ$("TEMP")
If LenB(m_TempFolder) = 0 Then
    m_TempFolder = App.Path
End If

m_RunId = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(CLng(Timer * 1000))
m_Assertions = 0
m_Failures = 0

Debug.Print "VB6FastIni regression tests"
Debug.Print String$(32, "-")

RunTest "WriteTextFile truncates shorter and empty content"
RunTest "Lookups are case-insensitive and formatting is preserved"
RunTest "SaveAs writes a copy and updates object state"
RunTest "Deleting the final section saves an empty file"
RunTest "DeleteKey removes only the requested key"

DeleteTestFile TestFileName("write")
DeleteTestFile TestFileName("format")
DeleteTestFile TestFileName("saveas_source")
DeleteTestFile TestFileName("saveas_target")
DeleteTestFile TestFileName("empty")
DeleteTestFile TestFileName("deletekey")

Debug.Print String$(32, "-")
Debug.Print CStr(m_Assertions) & " assertions; " & CStr(m_Failures) & " failures"

If m_Failures = 0 Then
    MsgBox CStr(m_Assertions) & " assertions passed.", vbInformation, "VB6FastIni Tests"
Else
    MsgBox CStr(m_Failures) & " failure(s) out of " & CStr(m_Assertions) & _
        " assertions. See the Immediate window for details.", vbExclamation, "VB6FastIni Tests"
End If

End Sub

Private Sub RunTest(ByVal TestName As String)
Dim FailuresBefore As Long

FailuresBefore = m_Failures
On Error GoTo TestError

Select Case TestName
    Case "WriteTextFile truncates shorter and empty content"
        TestWriteTextFileTruncates
    Case "Lookups are case-insensitive and formatting is preserved"
        TestCaseInsensitiveLookupAndFormatting
    Case "SaveAs writes a copy and updates object state"
        TestSaveAs
    Case "Deleting the final section saves an empty file"
        TestDeleteFinalSection
    Case "DeleteKey removes only the requested key"
        TestDeleteKey
End Select

On Error GoTo 0
If m_Failures = FailuresBefore Then
    Debug.Print "PASS  " & TestName
Else
    Debug.Print "FAIL  " & TestName
End If
Exit Sub

TestError:
RecordFailure TestName & " raised error " & CStr(Err.Number) & ": " & Err.Description
Resume TestFailed

TestFailed:
On Error GoTo 0
Debug.Print "FAIL  " & TestName

End Sub

Private Sub TestWriteTextFileTruncates()
Dim FileName As String

FileName = TestFileName("write")
WriteTextFile FileName, "This content is longer than the replacement."
WriteTextFile FileName, "Short"
AssertEqualString "shorter write", "Short", ReadTextFile(FileName)

WriteTextFile FileName, ""
AssertEqualLong "empty write file length", 0, FileLen(FileName)
AssertEqualString "empty write content", "", ReadTextFile(FileName)

End Sub

Private Sub TestCaseInsensitiveLookupAndFormatting()
Dim FileName As String
Dim Ini As New cFastIni
Dim OriginalText As String
Dim ExpectedText As String

FileName = TestFileName("format")
OriginalText = "; keep this comment" & vbCrLf & vbCrLf & _
    "[General]" & vbCrLf & "Name=Before" & vbCrLf
ExpectedText = "; keep this comment" & vbCrLf & vbCrLf & _
    "[General]" & vbCrLf & "Name=After" & vbCrLf

WriteTextFile FileName, OriginalText
Ini.LoadIni FileName

AssertEqualString "case-insensitive read", "Before", Ini.ReadString("general", "name")
Ini.WriteString "GENERAL", "NAME", "After"
AssertEqualBoolean "write marks object dirty", True, Ini.Dirty
Ini.Save

AssertEqualString "comments, blank line, and CRLF retained", ExpectedText, ReadTextFile(FileName)
AssertEqualBoolean "save clears dirty flag", False, Ini.Dirty

End Sub

Private Sub TestSaveAs()
Dim SourceFile As String
Dim TargetFile As String
Dim OriginalText As String
Dim ExpectedText As String
Dim Ini As New cFastIni

SourceFile = TestFileName("saveas_source")
TargetFile = TestFileName("saveas_target")
OriginalText = "[General]" & vbCrLf & "Name=Before"
ExpectedText = "[General]" & vbCrLf & "Name=After"

WriteTextFile SourceFile, OriginalText
Ini.LoadIni SourceFile
Ini.WriteString "General", "Name", "After"
Ini.SaveAs TargetFile

AssertEqualString "SaveAs target content", ExpectedText, ReadTextFile(TargetFile)
AssertEqualString "SaveAs updates file name", TargetFile, Ini.FileName
AssertEqualBoolean "SaveAs clears dirty flag", False, Ini.Dirty
AssertEqualString "SaveAs leaves source unchanged", OriginalText, ReadTextFile(SourceFile)

End Sub

Private Sub TestDeleteFinalSection()
Dim FileName As String
Dim Ini As New cFastIni

FileName = TestFileName("empty")
WriteTextFile FileName, "[General]" & vbCrLf & "Name=Before"
Ini.LoadIni FileName
Ini.DeleteSection "General"
Ini.Save

AssertEqualLong "final section deletion file length", 0, FileLen(FileName)
AssertEqualString "final section deletion content", "", ReadTextFile(FileName)

End Sub

Private Sub TestDeleteKey()
Dim FileName As String
Dim Ini As New cFastIni

FileName = TestFileName("deletekey")
WriteTextFile FileName, "[General]" & vbCrLf & "Name=Before" & vbCrLf & "Age=42"
Ini.LoadIni FileName
Ini.DeleteKey "general", "name"

AssertEqualBoolean "deleted key no longer exists", False, Ini.KeyExists("General", "Name")
AssertEqualString "other key remains", "42", Ini.ReadString("General", "Age")
Ini.Save
AssertEqualString "delete key saved", "[General]" & vbCrLf & "Age=42", ReadTextFile(FileName)

End Sub

Private Function TestFileName(ByVal Suffix As String) As String

TestFileName = m_TempFolder & "\VB6FastIniTests_" & m_RunId & "_" & Suffix & ".ini"

End Function

Private Sub DeleteTestFile(ByVal FileName As String)

If LenB(Dir$(FileName)) > 0 Then
    Kill FileName
End If

End Sub

Private Sub AssertEqualString(ByVal AssertionName As String, ByVal Expected As String, ByVal Actual As String)

m_Assertions = m_Assertions + 1
If StrComp(Expected, Actual, vbBinaryCompare) <> 0 Then
    RecordFailure AssertionName & ": expected [" & DisplayText(Expected) & _
        "], got [" & DisplayText(Actual) & "]"
End If

End Sub

Private Sub AssertEqualLong(ByVal AssertionName As String, ByVal Expected As Long, ByVal Actual As Long)

m_Assertions = m_Assertions + 1
If Expected <> Actual Then
    RecordFailure AssertionName & ": expected " & CStr(Expected) & ", got " & CStr(Actual)
End If

End Sub

Private Sub AssertEqualBoolean(ByVal AssertionName As String, ByVal Expected As Boolean, ByVal Actual As Boolean)

m_Assertions = m_Assertions + 1
If Expected <> Actual Then
    RecordFailure AssertionName & ": expected " & CStr(Expected) & ", got " & CStr(Actual)
End If

End Sub

Private Sub RecordFailure(ByVal Detail As String)

m_Failures = m_Failures + 1
Debug.Print "  " & Detail

End Sub

Private Function DisplayText(ByVal Text As String) As String

DisplayText = Replace(Text, vbCrLf, "<CRLF>")

End Function
