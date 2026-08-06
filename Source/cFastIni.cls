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
Private Function FindKey( _
    ByVal Section As String, _
    ByVal Key As String) As Long

    Dim i As Long

    For i = 0 To m_LineCount - 1

        If m_Lines(i).LineType = ilKey Then

            If StrComp(m_Lines(i).Section, Section, vbTextCompare) = 0 _
            And StrComp(m_Lines(i).Key, Key, vbTextCompare) = 0 Then

                FindKey = i
                Exit Function

            End If

        End If

    Next

    FindKey = -1

End Function
Public Function ReadString( _
    ByVal Section As String, _
    ByVal Key As String, _
    Optional ByVal DefaultValue As String = "") As String

    Dim Index As Long

    Index = FindKey(Section, Key)

    If Index >= 0 Then
        ReadString = m_Lines(Index).Value
    Else
        ReadString = DefaultValue
    End If

End Function
Private Sub Clear()

    m_LineCount = 0

    Erase m_Lines

End Sub
Public Property Get FileName() As String: FileName = m_FileName: End Property
Public Property Get Dirty() As Boolean: Dirty = m_Dirty: End Property

Public Sub Load(ByVal FileName As String)

Dim Text As String
Dim Parts() As String

Dim i As Long

Dim CurrentSection As String

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

            CurrentSection = Mid$(Trim$(Parts(i)), 2, _
                                  Len(Trim$(Parts(i))) - 2)

            m_Lines(i).LineType = ilSection

            m_Lines(i).Section = CurrentSection

        Case Else

            Dim K As String
            Dim V As String

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
Public Sub Save()
    If LenB(m_FileName) = 0 Then Err.Raise 5, "cFastIni.Save", "No file loaded."
    m_Dirty = False
End Sub
