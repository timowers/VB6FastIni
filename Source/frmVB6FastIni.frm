VERSION 5.00
Begin VB.Form frmVB6FastIni 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmVB6FastIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim Ini As New cFastIni

Ini.Load App.Path & "\Test.ini"

Debug.Print Ini.ReadString("General", "Name")
Debug.Print Ini.ReadString("Display", "Width")
Debug.Print Ini.ReadString("Display", "Missing", "Default")

End

End Sub
