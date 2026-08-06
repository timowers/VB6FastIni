VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cFastIni"
Option Explicit

Private Type TEntry
    Section As String
    Key As String
    Value As String
End Type

Private m_FileName As String
Private m_Dirty As Boolean
Private m_Items() As TEntry
Private m_Count As Long

Public Property Get FileName() As String: FileName = m_FileName: End Property
Public Property Get Dirty() As Boolean: Dirty = m_Dirty: End Property

Public Sub Load(ByVal FileName As String)
    m_FileName = FileName
    m_Count = 0
    ReDim m_Items(0)
    m_Dirty = False
End Sub

Public Sub Save()
    If LenB(m_FileName)=0 Then Err.Raise 5,"cFastIni.Save","No file loaded."
    m_Dirty=False
End Sub

Private Sub EnsureCapacity()
    If m_Count>UBound(m_Items) Then ReDim Preserve m_Items((UBound(m_Items)+1)*2+1)
End Sub

Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String)
    EnsureCapacity
    m_Items(m_Count).Section=Section
    m_Items(m_Count).Key=Key
    m_Items(m_Count).Value=Value
    m_Count=m_Count+1
    m_Dirty=True
End Sub
