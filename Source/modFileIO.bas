Attribute VB_Name = "modFileIO"
Option Explicit
Public Function ReadTextFile(ByVal FileName As String, Optional ByRef IsUTF8 As Boolean = False) As String
Dim FileNumber As Integer
Dim FileIsOpen As Boolean
Dim ErrorNumber As Long
Dim ErrorDescription As String
Dim FileBytes() As Byte
Dim FileSize As Long
Dim LastByte As Long

On Error GoTo ErrorHandler

IsUTF8 = False
FileNumber = FreeFile

Open FileName For Binary Access Read As #FileNumber
FileIsOpen = True

FileSize = LOF(FileNumber)
If FileSize > 0 Then
    ReDim FileBytes(0 To FileSize - 1)
    Get #FileNumber, , FileBytes
End If

Close #FileNumber
FileIsOpen = False

If FileSize = 0 Then Exit Function
LastByte = FileSize - 1

If FileSize >= 4 Then
    If (FileBytes(0) = &HFF And FileBytes(1) = &HFE And _
        FileBytes(2) = 0 And FileBytes(3) = 0) Or _
       (FileBytes(0) = 0 And FileBytes(1) = 0 And _
        FileBytes(2) = &HFE And FileBytes(3) = &HFF) Then
        Err.Raise 5, "ReadTextFile", "UTF-32 INI files are not supported."
    End If
End If

If FileSize >= 2 Then
    If (FileBytes(0) = &HFF And FileBytes(1) = &HFE) Or _
       (FileBytes(0) = &HFE And FileBytes(1) = &HFF) Then
        Err.Raise 5, "ReadTextFile", "UTF-16 INI files are not supported."
    End If
End If

If FileSize >= 3 Then
    If FileBytes(0) = &HEF And FileBytes(1) = &HBB And FileBytes(2) = &HBF Then
        IsUTF8 = True
        ReadTextFile = DecodeUTF8(FileBytes, 3)
        Exit Function
    End If
End If

ReadTextFile = StrConv(FileBytes, vbUnicode)
Exit Function

ErrorHandler:

ErrorNumber = Err.Number
ErrorDescription = Err.Description
On Error Resume Next
If FileIsOpen Then Close #FileNumber
On Error GoTo 0

Err.Raise ErrorNumber, "ReadTextFile", ErrorDescription

End Function

Public Sub WriteTextFile(ByVal FileName As String, ByVal Text As String, Optional ByVal IsUTF8 As Boolean = False)
Dim FileNumber As Integer
Dim FileIsOpen As Boolean
Dim ErrorNumber As Long
Dim ErrorDescription As String
Dim FileBytes() As Byte

On Error GoTo ErrorHandler

If IsUTF8 Then EncodeUTF8 Text, FileBytes
FileNumber = FreeFile

If IsUTF8 Then
    ' Truncate with Output mode, then write the UTF-8 byte stream in Binary mode.
    Open FileName For Output As #FileNumber
    FileIsOpen = True
    Close #FileNumber
    FileIsOpen = False

    FileNumber = FreeFile
    Open FileName For Binary Access Write As #FileNumber
    FileIsOpen = True
    Put #FileNumber, , FileBytes
Else
    Open FileName For Output As #FileNumber
    FileIsOpen = True
    Print #FileNumber, Text;
End If

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

Private Function DecodeUTF8(ByRef FileBytes() As Byte, ByVal FirstByte As Long) As String
Dim TextParts() As String
Dim ByteIndex As Long
Dim LastByte As Long
Dim PartIndex As Long
Dim Lead As Long
Dim Second As Long
Dim Third As Long
Dim Fourth As Long
Dim CodePoint As Long

LastByte = UBound(FileBytes)
If FirstByte > LastByte Then Exit Function
ReDim TextParts(0 To LastByte - FirstByte)
ByteIndex = FirstByte

Do While ByteIndex <= LastByte
    Lead = CLng(FileBytes(ByteIndex))

    If Lead <= &H7F Then
        CodePoint = Lead
        ByteIndex = ByteIndex + 1
    ElseIf Lead >= &HC2 And Lead <= &HDF Then
        If ByteIndex + 1 > LastByte Then GoTo InvalidUTF8
        Second = CLng(FileBytes(ByteIndex + 1))
        If Second < &H80 Or Second > &HBF Then GoTo InvalidUTF8
        CodePoint = (Lead And &H1F) * &H40 + (Second And &H3F)
        ByteIndex = ByteIndex + 2
    ElseIf Lead >= &HE0 And Lead <= &HEF Then
        If ByteIndex + 2 > LastByte Then GoTo InvalidUTF8
        Second = CLng(FileBytes(ByteIndex + 1))
        Third = CLng(FileBytes(ByteIndex + 2))
        If Second < &H80 Or Second > &HBF Or Third < &H80 Or Third > &HBF Then GoTo InvalidUTF8
        If Lead = &HE0 And Second < &HA0 Then GoTo InvalidUTF8
        If Lead = &HED And Second > &H9F Then GoTo InvalidUTF8
        CodePoint = (Lead And &HF) * &H1000 + (Second And &H3F) * &H40 + (Third And &H3F)
        ByteIndex = ByteIndex + 3
    ElseIf Lead >= &HF0 And Lead <= &HF4 Then
        If ByteIndex + 3 > LastByte Then GoTo InvalidUTF8
        Second = CLng(FileBytes(ByteIndex + 1))
        Third = CLng(FileBytes(ByteIndex + 2))
        Fourth = CLng(FileBytes(ByteIndex + 3))
        If Second < &H80 Or Second > &HBF Or Third < &H80 Or Third > &HBF Or _
            Fourth < &H80 Or Fourth > &HBF Then GoTo InvalidUTF8
        If Lead = &HF0 And Second < &H90 Then GoTo InvalidUTF8
        If Lead = &HF4 And Second > &H8F Then GoTo InvalidUTF8
        CodePoint = (Lead And 7) * &H40000 + (Second And &H3F) * &H1000 + _
            (Third And &H3F) * &H40 + (Fourth And &H3F)
        ByteIndex = ByteIndex + 4
    Else
        GoTo InvalidUTF8
    End If

    If CodePoint <= &HFFFF Then
        TextParts(PartIndex) = UnicodeCharacter(CodePoint)
    Else
        CodePoint = CodePoint - &H10000
        TextParts(PartIndex) = UnicodeCharacter(&HD800 + (CodePoint \ &H400)) & _
            UnicodeCharacter(&HDC00 + (CodePoint And &H3FF))
    End If
    PartIndex = PartIndex + 1
Loop

DecodeUTF8 = Join(TextParts, vbNullString)
Exit Function

InvalidUTF8:
Err.Raise 13, "ReadTextFile", "The file contains invalid UTF-8 data."

End Function

Private Sub EncodeUTF8(ByVal Text As String, ByRef FileBytes() As Byte)
Dim TextIndex As Long
Dim ByteIndex As Long
Dim ByteCount As Long
Dim CodePoint As Long

ByteCount = 3
TextIndex = 1
Do While TextIndex <= Len(Text)
    CodePoint = NextCodePoint(Text, TextIndex)
    If CodePoint <= &H7F Then
        ByteCount = ByteCount + 1
    ElseIf CodePoint <= &H7FF Then
        ByteCount = ByteCount + 2
    ElseIf CodePoint <= &HFFFF Then
        ByteCount = ByteCount + 3
    Else
        ByteCount = ByteCount + 4
    End If
    TextIndex = TextIndex + 1
Loop

ReDim FileBytes(0 To ByteCount - 1)
FileBytes(0) = &HEF
FileBytes(1) = &HBB
FileBytes(2) = &HBF
ByteIndex = 3
TextIndex = 1

Do While TextIndex <= Len(Text)
    CodePoint = NextCodePoint(Text, TextIndex)
    If CodePoint <= &H7F Then
        FileBytes(ByteIndex) = CodePoint
        ByteIndex = ByteIndex + 1
    ElseIf CodePoint <= &H7FF Then
        FileBytes(ByteIndex) = &HC0 Or (CodePoint \ &H40)
        FileBytes(ByteIndex + 1) = &H80 Or (CodePoint And &H3F)
        ByteIndex = ByteIndex + 2
    ElseIf CodePoint <= &HFFFF Then
        FileBytes(ByteIndex) = &HE0 Or (CodePoint \ &H1000)
        FileBytes(ByteIndex + 1) = &H80 Or ((CodePoint \ &H40) And &H3F)
        FileBytes(ByteIndex + 2) = &H80 Or (CodePoint And &H3F)
        ByteIndex = ByteIndex + 3
    Else
        FileBytes(ByteIndex) = &HF0 Or (CodePoint \ &H40000)
        FileBytes(ByteIndex + 1) = &H80 Or ((CodePoint \ &H1000) And &H3F)
        FileBytes(ByteIndex + 2) = &H80 Or ((CodePoint \ &H40) And &H3F)
        FileBytes(ByteIndex + 3) = &H80 Or (CodePoint And &H3F)
        ByteIndex = ByteIndex + 4
    End If
    TextIndex = TextIndex + 1
Loop

End Sub

Private Function NextCodePoint(ByVal Text As String, ByRef TextIndex As Long) As Long
Dim HighUnit As Long
Dim LowUnit As Long

HighUnit = AscW(Mid$(Text, TextIndex, 1))
If HighUnit < 0 Then HighUnit = HighUnit + &H10000

If HighUnit >= &HD800 And HighUnit <= &HDBFF Then
    If TextIndex >= Len(Text) Then GoTo InvalidUnicode
    LowUnit = AscW(Mid$(Text, TextIndex + 1, 1))
    If LowUnit < 0 Then LowUnit = LowUnit + &H10000
    If LowUnit < &HDC00 Or LowUnit > &HDFFF Then GoTo InvalidUnicode
    NextCodePoint = &H10000 + (HighUnit - &HD800) * &H400 + (LowUnit - &HDC00)
    TextIndex = TextIndex + 1
ElseIf HighUnit >= &HDC00 And HighUnit <= &HDFFF Then
    GoTo InvalidUnicode
Else
    NextCodePoint = HighUnit
End If

Exit Function

InvalidUnicode:
Err.Raise 13, "WriteTextFile", "The string contains an unpaired UTF-16 surrogate."

End Function

Private Function UnicodeCharacter(ByVal CodeUnit As Long) As String

If CodeUnit > &H7FFF Then CodeUnit = CodeUnit - &H10000
UnicodeCharacter = ChrW$(CodeUnit)

End Function
