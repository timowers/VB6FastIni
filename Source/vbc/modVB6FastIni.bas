Attribute VB_Name = "modVB6FastIni"
Option Explicit

'==============
' Public Enums
'==============
    Public Enum IniLineType
        ilBlank = 0
        ilComment = 1
        ilSection = 2
        ilKey = 3
    End Enum

'==================
' Public Type Defs
'==================
    Public Type TIniLine
        LineType As IniLineType
        OriginalText As String
        Section As String
        Key As String
        Value As String
        Modified As Boolean
    End Type
Public Function IsComment(ByVal S As String) As Boolean

If Len(S) = 0 Then
    Exit Function
End If

Select Case Left$(S, 1)
    Case ";", "#"
        IsComment = True
End Select

End Function
Public Function IsSection(ByVal S As String) As Boolean

If Len(S) < 3 Then
    Exit Function
End If

IsSection = (Left$(S, 1) = "[") And (Right$(S, 1) = "]")

End Function
Public Function SplitKeyValue(ByVal S As String, ByRef Key As String, ByRef Value As String) As Boolean
Dim EqualsPos As Long

EqualsPos = InStr(1, S, "=")

If EqualsPos = 0 Then
    Exit Function
End If

Key = Trim$(Left$(S, EqualsPos - 1))
Value = Mid$(S, EqualsPos + 1)

SplitKeyValue = True

End Function

Public Function CompareText(ByVal LeftText As String, ByVal RightText As String) As Boolean

CompareText = (StrComp(LeftText, RightText, vbTextCompare) = 0)

End Function

