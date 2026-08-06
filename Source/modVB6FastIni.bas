Attribute VB_Name = "modFastIni"
Option Explicit

Public Enum IniLineType

    ilBlank = 0
    ilComment = 1
    ilSection = 2
    ilKey = 3

End Enum

Public Type TIniLine

    LineType As IniLineType

    OriginalText As String

    Section As String

    Key As String

    Value As String

    Modified As Boolean

End Type

Public Function IsComment(ByVal S As String) As Boolean

    If Len(S) = 0 Then Exit Function

    Select Case Left$(S, 1)

        Case ";", "#"

            IsComment = True

    End Select

End Function

Public Function IsSection(ByVal S As String) As Boolean

    If Len(S) < 3 Then Exit Function

    IsSection = _
        Left$(S, 1) = "[" _
    And Right$(S, 1) = "]"

End Function

Public Function SplitKeyValue(ByVal S As String, ByRef Key As String, ByRef Value As String) As Boolean

Dim lEqualsPos As Long

lEqualsPos = InStr(1, S, "=")

If lEqualsPos = 0 Then Exit Function

Key = Trim$(Left$(S, lEqualsPos - 1))
Value = Mid$(S, lEqualsPos + 1)

SplitKeyValue = True

End Function
Public Function CompareText(ByVal LeftText As String, ByVal RightText As String) As Boolean

    CompareText = (StrComp(LeftText, RightText, vbTextCompare) = 0)

End Function
