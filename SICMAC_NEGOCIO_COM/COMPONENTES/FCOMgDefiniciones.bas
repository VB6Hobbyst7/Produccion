Attribute VB_Name = "FCOMgDefiniciones"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public unidad(0 To 9) As String
Public decena(0 To 9) As String
Public centena(0 To 10) As String
Public deci(0 To 9) As String
Public otros(0 To 15) As String


Public cad As String
Public Cadd As String

'******************************************************************************
Global gsSimbolo As String
Global Const gcMN = "S/."
Global Const gcME = "$"

'*****************************************************************************
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'Public Const CB_SETDROPPEDWIDTH = &H160
'Public Const CB_FINDSTRING = &H14C
'Global Const gsConnServDBF = "DSN=DSNCmactServ"
Global gsCentralPers As String

