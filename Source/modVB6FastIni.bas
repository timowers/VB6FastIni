Attribute VB_Name = "modVB6FastIni"
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

    If