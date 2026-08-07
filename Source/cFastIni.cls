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

Private m_FileName As String
Private m_Dirty As Boolean
Private m_Count As Long
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
Public Function ReadString( _
    ByVal Section As String, _
    ByVal Key As String, _
    Optional ByVal DefaultValue As String = "") As String

On Error GoTo ErrorHandler

    Dim Index As Long

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
Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String)

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
        Err.Raise vbObjectError + 1000, "cFastIni.WriteString", "Section does not exist."
    End If

    LineIndex = SectionIndex + 1

    Do While LineIndex < m_LineCount

        If m_Lines(LineIndex).LineType = ilSection Then Exit Do

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
Private Function BuildKeyLine(ByVal LineIndex As Long) As String

    BuildKeyLine = m_Lines(LineIndex).Key & "=" & m_Lines(LineIndex).Value

End Function
Private Sub Clear()

    m_LineCount = 0
    m_LineCapacity = 0

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
Public Property Get FileName() As String: FileName = m_FileName: End Property

Public Sub Load(ByVal FileName As String)
Dim Text As String
Dim Parts() As String
Dim i As Long
Dim CurrentSection As String
Dim K As String
Dim V As String

Clear

m_FileName = FileName

Text = ReadTextFile(FileName)

If Len(Text) = 0 Then Exit Sub

Parts = Split(Replace(Text, vbCrLf, vbLf), vbLf)

ReDim m_Lines(UBound(Parts))

For i = 0 To UBound(Parts)

    m_Lines(i).OriginalText = Parts(i)

    Select Case True

        Case Len(Trim$(Parts(i))) = 0

            m_Lines(i).LineType = ilBlank

        Case IsComment(Trim$(Parts(i)))

            m_Lines(i).LineType = ilComment

        Case IsSection(Trim$(Parts(i)))

            CurrentSection = Mid$(Trim$(Parts(i)), 2, Len(Trim$(Parts(i))) - 2)

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
Dim FileText As String
Dim LineIndex As Long

If LenB(m_FileName) = 0 Then
    Err.Raise 5, "cFastIni.Save", "No file loaded."
End If

If Not m_Dirty Then
    'm_Dirty will become m_IsModified
    Exit Sub
End If

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

WriteTextFile m_FileName, FileText

m_Dirty = False

End Sub
Private Sub EnsureLineCapacity(ByVal RequiredCount As Long)
Dim NewCapacity As Long

If RequiredCount <= m_LineCapacity Then Exit Sub

NewCapacity = m_LineCapacity

If NewCapacity = 0 Then NewCapacity = 16

Do While NewCapacity < RequiredCount
    NewCapacity = NewCapacity * 2
Loop

ReDim Preserve m_Lines(NewCapacity - 1)

m_LineCapacity = NewCapacity

End Sub
