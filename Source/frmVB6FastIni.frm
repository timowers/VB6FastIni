VERSION 5.00
Begin VB.Form frmFastIni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmFastIni"
   ClientHeight    =   2316
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3624
   Icon            =   "frmVB6FastIni.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
End
Attribute VB_Name = "frmFastIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==================
' Public Variables
'==================
    Public Ini As New cFastIni

Private Sub cmdTest_Click()
Dim OutputFile As String

On Error GoTo TestError

Debug.Print "General section exists: "; Ini.SectionExists("General")
Debug.Print "Name key exists: "; Ini.KeyExists("General", "Name")
Debug.Print "Name: "; Ini.ReadString("General", "Name", "<missing>")

Ini.WriteString "General", "Name", "Fred"
Debug.Print "Updated name in memory: "; Ini.ReadString("General", "Name")

OutputFile = Environ$("TEMP")
If LenB(OutputFile) = 0 Then OutputFile = App.Path
OutputFile = OutputFile & "\VB6FastIni-demo.ini"
Ini.SaveAs OutputFile

Debug.Print "Saved modified copy to: "; OutputFile
Exit Sub

TestError:
MsgBox "Error: " & Err.Description & vbNewLine & "Number: " & Err.Number, _
    vbExclamation, "VB6FastIni demo"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then
    Unload Me
End If

End Sub

Private Sub Form_Load()

On Error GoTo Form_Load_Error

Ini.LoadIni App.Path & "\Test.ini"

Exit Sub

Form_Load_Error:

MsgBox "Error: " & Err.Description & vbNewLine & "Number: " & Err.Number & vbNewLine & "Line: " & Erl & vbNewLine & "Sub: Form_Load" & vbNewLine & "Form: frmFastIni", vbCritical, "Unexpected Error In " & App.FileDescription

Resume Next

End Sub

