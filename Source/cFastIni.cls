VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFastIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===================
' Private Variables
'===================
    Private m_FileName As String
    Private m_Dirty As Boolean
    Private m_Lines() As TIniLine
    Private m_LineCount As Long
    Private m_LineCapacity As Long

Public Property Get Dirty() As Boolean

Dirty = m_Dirty

End Property
Private Function FindKey(ByVal Section As String, ByVal Key As String) As Long
Dim LineIndex As Long

For LineIndex = 0 To m_LineCount - 1
    If m_Lines(LineIndex).LineType = ilKey Then
        If CompareText(m_Lines(LineIndex).Section, Section) Then
            If CompareText(m_Lines(LineIndex).Key, Key) Then
                FindKey = LineIndex
                Exit Function
            End If
        End If
    End If
Next

FindKey = -1

End Function

Public Function ReadString(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As String = "") As String
Dim Index As Long

On Error GoTo ErrorHandler

Index = FindKey(Section, Key)

If Index >= 0 Then
    ReadString = m_Lines(Index).Value
Else
    ReadString = DefaultValue
End If

Exit Function

ErrorHandler:

Err.Raise Err.Number, "cFastIni.ReadString", Err.Description

End Function

Public Function ReadInteger(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As Long = 0) As Long
Dim Index As Long
Dim TextValue As String
Dim NumericValue As Double

Index = FindKey(Section, Key)

If Index < 0 Then
    ReadInteger = DefaultValue
    Exit Function
End If

TextValue = Trim$(m_Lines(Index).Value)

If Len(TextValue) = 0 Then
    Err.Raise 13, "cFastIni.ReadInteger", "Value for [" & Section & "] " & Key & " is empty."
End If

If Not IsNumeric(TextValue) Then
    Err.Raise 13, "cFastIni.ReadInteger", "Value for [" & Section & "] " & Key & " is not numeric."
End If

NumericValue = CDbl(TextValue)

If NumericValue <> Fix(NumericValue) Then
    Err.Raise 13, "cFastIni.ReadInteger", "Value for [" & Section & "] " & Key & " is not a whole number."
End If

ReadInteger = CLng(NumericValue)

End Function
Public Function ReadDouble(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As Double = 0#) As Double
Dim Index As Long
Dim TextValue As String

Index = FindKey(Section, Key)

If Index < 0 Then
    ReadDouble = DefaultValue
    Exit Function
End If

TextValue = Trim$(m_Lines(Index).Value)

If Len(TextValue) = 0 Then
    Err.Raise 13, "cFastIni.ReadDouble", "Value for [" & Section & "] " & Key & " is empty."
End If

If Not IsNumeric(TextValue) Then
    Err.Raise 13, "cFastIni.ReadDouble", "Value for [" & Section & "] " & Key & " is not numeric."
End If

ReadDouble = CDbl(TextValue)

End Function
Public Function ReadBoolean(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As Boolean = False) As Boolean
Dim TextValue As String

TextValue = LCase$(Trim$(ReadString(Section, Key, "")))

Select Case TextValue
    Case "true", "yes", "on", "1", "-1"
        ReadBoolean = True
    Case "false", "no", "off", "0"
        ReadBoolean = False
    Case Else
        ReadBoolean = DefaultValue
End Select

End Function

Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String)
'Modifies memory, not the ini file. Save > synchronises memory with disk
Dim LineIndex As Long
Dim SectionIndex As Long

LineIndex = FindKey(Section, Key)

If LineIndex >= 0 Then
    If m_Lines(LineIndex).Value <> Value Then
        m_Lines(LineIndex).Value = Value
        m_Lines(LineIndex).Modified = True
        m_Dirty = True
    End If
    Exit Sub
End If

SectionIndex = FindSection(Section)

If SectionIndex < 0 Then
    AddSection Section
    AddKey Section, Key, Value
    m_Dirty = True
    Exit Sub
End If

LineIndex = SectionIndex + 1

Do While LineIndex < m_LineCount
    If m_Lines(LineIndex).LineType = ilSection Then
        Exit Do
    End If
    LineIndex = LineIndex + 1
Loop

InsertLine LineIndex

m_Lines(LineIndex).LineType = ilKey
m_Lines(LineIndex).Section = Section
m_Lines(LineIndex).Key = Key
m_Lines(LineIndex).Value = Value
m_Lines(LineIndex).OriginalText = Key & "=" & Value
m_Lines(LineIndex).Modified = True

m_Dirty = True

End Sub
Public Sub DeleteKey(ByVal Section As String, ByVal Key As String)
Dim LineIndex As Long
Dim Index As Long

LineIndex = FindKey(Section, Key)

If LineIndex < 0 Then Exit Sub

For Index = LineIndex To m_LineCount - 2
    m_Lines(Index) = m_Lines(Index + 1)
Next Index

m_LineCount = m_LineCount - 1

m_Dirty = True

End Sub
Public Sub DeleteSection(ByVal Section As String)
Dim SectionIndex As Long
Dim EndIndex As Long
Dim RemoveCount As Long
Dim Index As Long

SectionIndex = FindSection(Section)

If SectionIndex < 0 Then
    Exit Sub
End If

EndIndex = SectionIndex + 1

Do While EndIndex < m_LineCount
    If m_Lines(EndIndex).LineType = ilSection Then
        Exit Do
    End If
    EndIndex = EndIndex + 1
Loop

RemoveCount = EndIndex - SectionIndex

For Index = SectionIndex To m_LineCount - RemoveCount - 1
    m_Lines(Index) = m_Lines(Index + RemoveCount)
Next Index

m_LineCount = m_LineCount - RemoveCount

m_Dirty = True

End Sub
Public Sub WriteInteger(ByVal Section As String, ByVal Key As String, ByVal Value As Long)

WriteString Section, Key, CStr(Value)

End Sub
Public Sub WriteDouble(ByVal Section As String, ByVal Key As String, ByVal Value As Double)

    WriteString Section, Key, CStr(Value)

End Sub
Public Sub WriteBoolean(ByVal Section As String, ByVal Key As String, ByVal Value As Boolean)

If Value Then
    WriteString Section, Key, "True"
Else
    WriteString Section, Key, "False"
End If

End Sub
Private Sub AddSection(ByVal Section As String)
Dim LineIndex As Long

EnsureLineCapacity m_LineCount + 1

LineIndex = m_LineCount

m_Lines(LineIndex).LineType = ilSection
m_Lines(LineIndex).Section = Section
m_Lines(LineIndex).OriginalText = "[" & Section & "]"

m_LineCount = m_LineCount + 1

End Sub
Private Sub AddKey(ByVal Section As String, ByVal Key As String, ByVal Value As String)
Dim LineIndex As Long

EnsureLineCapacity m_LineCount + 1

LineIndex = m_LineCount

m_Lines(LineIndex).LineType = ilKey
m_Lines(LineIndex).Section = Section
m_Lines(LineIndex).Key = Key
m_Lines(LineIndex).Value = Value
m_Lines(LineIndex).OriginalText = Key & "=" & Value
m_Lines(LineIndex).Modified = True

m_LineCount = m_LineCount + 1

End Sub
Private Function BuildKeyLine(ByVal LineIndex As Long) As String

    BuildKeyLine = m_Lines(LineIndex).Key & "=" & m_Lines(LineIndex).Value

End Function
Private Sub Clear()

m_LineCount = 0
m_LineCapacity = 0
m_Dirty = False
Erase m_Lines

End Sub
Private Sub InsertLine(ByVal LineIndex As Long)
Dim Index As Long

EnsureLineCapacity m_LineCount + 1

For Index = m_LineCount To LineIndex + 1 Step -1
    m_Lines(Index) = m_Lines(Index - 1)
Next Index

m_LineCount = m_LineCount + 1

End Sub
Public Property Get FileName() As String

    FileName = m_FileName

End Property

Public Sub LoadIni(ByVal FileName As String)
Dim Text As String
Dim Parts() As String
Dim i As Long
Dim CurrentSection As String
Dim K As String
Dim V As String

Text = ReadTextFile(FileName)

Clear

m_FileName = FileName

If Len(Text) = 0 Then
    Exit Sub
End If

Parts = Split(Replace(Text, vbCrLf, vbLf), vbLf)

ReDim m_Lines(UBound(Parts))

For i = 0 To UBound(Parts)
    m_Lines(i).OriginalText = Parts(i)
    Select Case True
        Case Len(Trim$(Parts(i))) = 0
            m_Lines(i).LineType = ilBlank
        Case modVB6FastIni.IsComment(Trim$(Parts(i)))
            m_Lines(i).LineType = ilComment
        Case IsSection(Trim$(Parts(i)))
            CurrentSection = Trim$(Mid$(Trim$(Parts(i)), 2, Len(Trim$(Parts(i))) - 2))
            m_Lines(i).LineType = ilSection
            m_Lines(i).Section = CurrentSection
        Case Else
            If SplitKeyValue(Parts(i), K, V) Then
                m_Lines(i).LineType = ilKey
                m_Lines(i).Section = CurrentSection
                m_Lines(i).Key = K
                m_Lines(i).Value = V
            Else
                m_Lines(i).LineType = ilComment
            End If
    End Select
Next

m_LineCount = UBound(m_Lines) + 1

End Sub

Private Function FindSection(ByVal Section As String) As Long
Dim LineIndex As Long

For LineIndex = 0 To m_LineCount - 1
    If m_Lines(LineIndex).LineType = ilSection Then
        If StrComp(m_Lines(LineIndex).Section, Section, vbTextCompare) = 0 Then
            FindSection = LineIndex
            Exit Function
        End If
    End If
Next

FindSection = -1

End Function

Public Function SectionExists(ByVal Section As String) As Boolean

SectionExists = (FindSection(Section) >= 0)

End Function

Public Function KeyExists(ByVal Section As String, ByVal Key As String) As Boolean

KeyExists = (FindKey(Section, Key) >= 0)

End Function

Public Sub Save()
If LenB(m_FileName) = 0 Then
    Err.Raise 5, "cFastIni.Save", "No file loaded."
End If

If Not m_Dirty Then
    Exit Sub
End If

WriteTextFile m_FileName, Serialize()

m_Dirty = False

End Sub
Public Sub SaveAs(ByVal FileName As String)

If LenB(FileName) = 0 Then
    Err.Raise 5, "cFastIni.SaveAs", "File name cannot be empty."
End If

WriteTextFile FileName, Serialize()

m_FileName = FileName
m_Dirty = False

End Sub
Private Function Serialize() As String
Dim FileText As String
Dim LineIndex As Long

For LineIndex = 0 To m_LineCount - 1
    If m_Lines(LineIndex).LineType = ilKey Then
        If m_Lines(LineIndex).Modified Then
            FileText = FileText & BuildKeyLine(LineIndex)
        Else
            FileText = FileText & m_Lines(LineIndex).OriginalText
        End If
    Else
        FileText = FileText & m_Lines(LineIndex).OriginalText
    End If
    If LineIndex < m_LineCount - 1 Then
        FileText = FileText & vbCrLf
    End If
Next

Serialize = FileText

End Function

Private Sub EnsureLineCapacity(ByVal RequiredCount As Long)
Dim NewCapacity As Long

If RequiredCount <= m_LineCapacity Then
    Exit Sub
End If

NewCapacity = m_LineCapacity

If NewCapacity = 0 Then
    NewCapacity = 16
End If

Do While NewCapacity < RequiredCount
    NewCapacity = NewCapacity * 2
Loop

ReDim Preserve m_Lines(NewCapacity - 1)

m_LineCapacity = NewCapacity

End Sub

