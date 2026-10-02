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
RunTest "LoadIni resets dirty state for non-empty and empty files"
RunTest "LoadIni missing file raises error and preserves loaded state"
RunTest "UTF-8 files round-trip through Save and SaveAsUTF8"
RunTest "Numeric reads default only for missing keys and reject invalid values"
RunTest "Boolean reads default only for missing keys and reject invalid values"

Close

DeleteTestFile TestFileName("write")
DeleteTestFile TestFileName("format")
DeleteTestFile TestFileName("saveas_source")
DeleteTestFile TestFileName("saveas_target")
DeleteTestFile TestFileName("empty")
DeleteTestFile TestFileName("deletekey")
DeleteTestFile TestFileName("load_first")
DeleteTestFile TestFileName("load_second")
DeleteTestFile TestFileName("load_empty")
DeleteTestFile TestFileName("load_missing_source")
DeleteTestFile TestFileName("load_missing")
DeleteTestFile TestFileName("utf8")
DeleteTestFile TestFileName("utf8_copy")
DeleteTestFile TestFileName("numeric")
DeleteTestFile TestFileName("boolean")

Debug.Print String$(32, "-")
Debug.Print CStr(m_Assertions) & " assertions; " & CStr(m_Failures) & " failures"

If m_Failures = 0 Then
    MsgBox CStr(m_Assertions) & " assertions passed.", vbInformation, "VB6FastIni Tests"
Else
    If m_Failures = 1 Then
        MsgBox "1 failure out of " & CStr(m_Assertions) & " assertions." & vbNewLine & "See the Immediate window for details.", vbExclamation, "VB6FastIni Tests"
    Else
        MsgBox CStr(m_Failures) & " failures out of " & CStr(m_Assertions) & " assertions." & vbNewLine & "See the Immediate window for details.", vbExclamation, "VB6FastIni Tests"
    End If
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
    Case "LoadIni resets dirty state for non-empty and empty files"
        TestLoadClearsDirtyState
    Case "LoadIni missing file raises error and preserves loaded state"
        TestLoadMissingFile
    Case "UTF-8 files round-trip through Save and SaveAsUTF8"
        TestUTF8RoundTrip
    Case "Numeric reads default only for missing keys and reject invalid values"
        TestNumericReadErrors
    Case "Boolean reads default only for missing keys and reject invalid values"
        TestBooleanReadErrors
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

Private Sub TestLoadClearsDirtyState()
Dim FirstFile As String
Dim SecondFile As String
Dim EmptyFile As String
Dim Ini As New cFastIni

FirstFile = TestFileName("load_first")
SecondFile = TestFileName("load_second")
EmptyFile = TestFileName("load_empty")

WriteTextFile FirstFile, "[First]" & vbCrLf & "Value=1"
WriteTextFile SecondFile, "[Second]" & vbCrLf & "Value=2"
WriteTextFile EmptyFile, ""

Ini.LoadIni FirstFile
Ini.WriteString "First", "Unsaved", "change"
AssertEqualBoolean "edit marks object dirty before reload", True, Ini.Dirty

Ini.LoadIni SecondFile
AssertEqualBoolean "non-empty load clears dirty flag", False, Ini.Dirty
AssertEqualString "non-empty load replaces data", "2", Ini.ReadString("Second", "Value")
AssertEqualString "non-empty load updates file name", SecondFile, Ini.FileName

Ini.WriteString "Second", "Unsaved", "change"
Ini.LoadIni EmptyFile
AssertEqualBoolean "empty load clears dirty flag", False, Ini.Dirty
AssertEqualBoolean "empty load removes previous sections", False, Ini.SectionExists("Second")
AssertEqualString "empty load updates file name", EmptyFile, Ini.FileName

End Sub

Private Sub TestLoadMissingFile()
Dim SourceFile As String
Dim MissingFile As String
Dim Ini As New cFastIni

SourceFile = TestFileName("load_missing_source")
MissingFile = TestFileName("load_missing")
WriteTextFile SourceFile, "[Existing]" & vbCrLf & "Value=kept"
Ini.LoadIni SourceFile

AssertEqualLong "missing LoadIni raises file-not-found", 53, LoadIniErrorNumber(Ini, MissingFile)
AssertEqualString "failed load preserves existing value", "kept", Ini.ReadString("Existing", "Value")
AssertEqualString "failed load preserves current file name", SourceFile, Ini.FileName
AssertEqualBoolean "failed load preserves clean state", False, Ini.Dirty

End Sub

Private Sub TestUTF8RoundTrip()
Dim FileName As String
Dim CopyFileName As String
Dim UnicodeText As String
Dim Ini As New cFastIni
Dim LoadedIni As New cFastIni
Dim CopyIni As New cFastIni

FileName = TestFileName("utf8")
CopyFileName = TestFileName("utf8_copy")
UnicodeText = "Cafe " & ChrW$(&HE9) & " - " & ChrW$(&H6771) & ChrW$(&H4EAC)

Ini.WriteString "Text", "Greeting", UnicodeText
Ini.SaveAsUTF8 FileName
AssertEqualString "UTF-8 read after SaveAsUTF8", UnicodeText, Ini.ReadString("Text", "Greeting")

Ini.WriteString "Text", "Greeting", UnicodeText & "!"
Ini.Save
LoadedIni.LoadIni FileName
AssertEqualString "UTF-8 read after Save", UnicodeText & "!", LoadedIni.ReadString("Text", "Greeting")

LoadedIni.SaveAs CopyFileName
CopyIni.LoadIni CopyFileName
AssertEqualString "SaveAs preserves UTF-8 encoding", UnicodeText & "!", CopyIni.ReadString("Text", "Greeting")

End Sub

Private Sub TestNumericReadErrors()
Dim FileName As String
Dim Ini As New cFastIni

FileName = TestFileName("numeric")
WriteTextFile FileName, "[Numbers]" & vbCrLf & _
    "IntegerValid=42" & vbCrLf & _
    "IntegerInvalid=abc" & vbCrLf & _
    "IntegerFraction=" & CStr(42.5) & vbCrLf & _
    "IntegerBlank=" & vbCrLf & _
    "IntegerOverflow=2147483648" & vbCrLf & _
    "DoubleValid=" & CStr(12.5) & vbCrLf & _
    "DoubleInvalid=abc" & vbCrLf & _
    "DoubleBlank="

Ini.LoadIni FileName

'AssertEqualBoolean "valid boolean", 0, Ini.ReadBoolean("Numbers", "IntegerValid")

AssertEqualLong "valid integer", 42, Ini.ReadInteger("Numbers", "IntegerValid")
AssertEqualDouble "valid double", 12.5, Ini.ReadDouble("Numbers", "DoubleValid")
AssertEqualLong "missing integer returns default", 7, Ini.ReadInteger("Numbers", "MissingInteger", 7)
AssertEqualDouble "missing double returns default", 2.5, Ini.ReadDouble("Numbers", "MissingDouble", 2.5)
AssertEqualLong "invalid integer raises type mismatch", 13, ReadIntegerErrorNumber(Ini, "Numbers", "IntegerInvalid")
AssertEqualLong "fractional integer raises type mismatch", 13, ReadIntegerErrorNumber(Ini, "Numbers", "IntegerFraction")
AssertEqualLong "empty integer raises type mismatch", 13, ReadIntegerErrorNumber(Ini, "Numbers", "IntegerBlank")
AssertEqualLong "out-of-range integer raises overflow", 6, ReadIntegerErrorNumber(Ini, "Numbers", "IntegerOverflow")
AssertEqualLong "invalid double raises type mismatch", 13, ReadDoubleErrorNumber(Ini, "Numbers", "DoubleInvalid")
AssertEqualLong "empty double raises type mismatch", 13, ReadDoubleErrorNumber(Ini, "Numbers", "DoubleBlank")

End Sub

Private Sub TestBooleanReadErrors()
Dim FileName As String
Dim Ini As New cFastIni

FileName = TestFileName("boolean")
WriteTextFile FileName, "[Flags]" & vbCrLf & _
    "TrueWord=true" & vbCrLf & _
    "TrueYes=yes" & vbCrLf & _
    "TrueOn=on" & vbCrLf & _
    "TrueOne=1" & vbCrLf & _
    "TrueMinusOne=-1" & vbCrLf & _
    "FalseWord=false" & vbCrLf & _
    "FalseNo=no" & vbCrLf & _
    "FalseOff=off" & vbCrLf & _
    "FalseZero=0" & vbCrLf & _
    "Invalid=perhaps" & vbCrLf & _
    "Blank="

Ini.LoadIni FileName

AssertEqualBoolean "true token true", True, Ini.ReadBoolean("Flags", "TrueWord")
AssertEqualBoolean "yes token true", True, Ini.ReadBoolean("Flags", "TrueYes")
AssertEqualBoolean "on token true", True, Ini.ReadBoolean("Flags", "TrueOn")
AssertEqualBoolean "1 token true", True, Ini.ReadBoolean("Flags", "TrueOne")
AssertEqualBoolean "-1 token true", True, Ini.ReadBoolean("Flags", "TrueMinusOne")
AssertEqualBoolean "false token false", False, Ini.ReadBoolean("Flags", "FalseWord")
AssertEqualBoolean "no token false", False, Ini.ReadBoolean("Flags", "FalseNo")
AssertEqualBoolean "off token false", False, Ini.ReadBoolean("Flags", "FalseOff")
AssertEqualBoolean "0 token false", False, Ini.ReadBoolean("Flags", "FalseZero")
AssertEqualBoolean "missing Boolean returns true default", True, Ini.ReadBoolean("Flags", "MissingTrue", True)
AssertEqualBoolean "missing Boolean returns false default", False, Ini.ReadBoolean("Flags", "MissingFalse", False)
AssertEqualLong "invalid Boolean raises type mismatch", 13, ReadBooleanErrorNumber(Ini, "Flags", "Invalid")
AssertEqualLong "blank Boolean raises type mismatch", 13, ReadBooleanErrorNumber(Ini, "Flags", "Blank")

End Sub

Private Function TestFileName(ByVal Suffix As String) As String

TestFileName = m_TempFolder & "\VB6FastIniTests_" & m_RunId & "_" & Suffix & ".ini"

End Function

Private Sub DeleteTestFile(ByVal FileName As String)
Dim ErrorNumber As Long
Dim ErrorDescription As String

On Error GoTo DeleteError

If LenB(Dir$(FileName)) > 0 Then
    Kill FileName
End If

Exit Sub

DeleteError:

ErrorNumber = Err.Number
ErrorDescription = Err.Description
On Error GoTo 0
RecordFailure "Could not delete temporary test file [" & FileName & "] (error " & _
    CStr(ErrorNumber) & ": " & ErrorDescription & ")"

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

Private Sub AssertEqualDouble(ByVal AssertionName As String, ByVal Expected As Double, ByVal Actual As Double)

m_Assertions = m_Assertions + 1
If Abs(Expected - Actual) > 0.0000001 Then
    RecordFailure AssertionName & ": expected " & CStr(Expected) & ", got " & CStr(Actual)
End If

End Sub

Private Function ReadBooleanErrorNumber(ByVal Ini As cFastIni, ByVal Section As String, ByVal Key As String) As Long
Dim Value As Boolean

Err.Clear
On Error Resume Next
Value = Ini.ReadBoolean(Section, Key)
ReadBooleanErrorNumber = Err.Number
On Error GoTo 0

End Function

Private Function LoadIniErrorNumber(ByVal Ini As cFastIni, ByVal FileName As String) As Long

Err.Clear
On Error Resume Next
Ini.LoadIni FileName
LoadIniErrorNumber = Err.Number
On Error GoTo 0

End Function

Private Function ReadIntegerErrorNumber(ByVal Ini As cFastIni, ByVal Section As String, ByVal Key As String) As Long
Dim Value As Long

Err.Clear
On Error Resume Next
Value = Ini.ReadInteger(Section, Key)
ReadIntegerErrorNumber = Err.Number
On Error GoTo 0

End Function

Private Function ReadDoubleErrorNumber(ByVal Ini As cFastIni, ByVal Section As String, ByVal Key As String) As Long
Dim Value As Double

Err.Clear
On Error Resume Next
Value = Ini.ReadDouble(Section, Key)
ReadDoubleErrorNumber = Err.Number
On Error GoTo 0

End Function

Private Sub RecordFailure(ByVal Detail As String)

m_Failures = m_Failures + 1
Debug.Print "  " & Detail

End Sub

Private Function DisplayText(ByVal Text As String) As String

DisplayText = Replace(Text, vbCrLf, "<CRLF>")

End Function
