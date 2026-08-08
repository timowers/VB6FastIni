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

Debug.Print Ini.ReadDouble("Numbers", "Pi")
Debug.Print Ini.ReadDouble("Numbers", "Missing", 12.5)
Debug.Print Ini.ReadDouble("Numbers", "Invalid", 99.9)

Exit Sub

Debug.Print Ini.ReadBoolean("Options", "Enabled")
Debug.Print Ini.ReadBoolean("Options", "Logging")
Debug.Print Ini.ReadBoolean("Options", "Caching")
Debug.Print Ini.ReadBoolean("Options", "Tracing")
Debug.Print Ini.ReadBoolean("Options", "Unknown", True)
Debug.Print Ini.ReadBoolean("Options", "Missing", True)

Debug.Print Ini.ReadInteger("Display", "Width")
Debug.Print Ini.ReadInteger("Display", "Missing", 1024)
Debug.Print Ini.ReadInteger("Display", "Invalid", 123)

Debug.Print Ini.ReadString("General", "Name")
Debug.Print Ini.ReadString("Display", "Width")
Debug.Print Ini.ReadString("Display", "Missing", "Default")

Debug.Print Ini.SectionExists("General")
Debug.Print Ini.SectionExists("Printer")

Debug.Print Ini.KeyExists("General", "Name")
Debug.Print Ini.KeyExists("General", "Missing")

Debug.Print

Debug.Print "Before"

Debug.Print Ini.ReadString("General", "Name")

Ini.WriteString "General", "Name", "Fred"

Debug.Print

Debug.Print "After"

Debug.Print Ini.ReadString("General", "Name")
Debug.Print Ini.ReadString("General", "Age")

Ini.Save

Debug.Print Ini.ReadString("General", "Description", "Default")

Ini.WriteString "Network", "Host", "192.168.1.10"
Ini.Save

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

