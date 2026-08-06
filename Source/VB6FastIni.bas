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

Public Function SplitKeyValue( _
                ByVal S As String, _
                ByRef Key As String, _
                ByRef Value As String) As Boolean

    Dim P As Long

    P = InStr(S, "=")

    If P = 0 Then Exit Function

    Key = Trim$(Left$(S, P - 1))

    Value = Mid$(S, P + 1)

    SplitKeyValue = True

End Function

