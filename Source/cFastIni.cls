VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cFastIni"
Option Explicit

' VB6FastIni v0.1.0
' Initial class skeleton.

Private m_FileName As String
Private m_Dirty As Boolean

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Get Dirty() As Boolean
    Dirty = m_Dirty
End Property

Public Sub Load(ByVal FileName As String)
    m_FileName = FileName
    m_Dirty = False
End Sub

Public Sub Save()
    If LenB(m_FileName) = 0 Then Err.Raise 5, "cFastIni.Save", "No file loaded."
    m_Dirty = False
End Sub
