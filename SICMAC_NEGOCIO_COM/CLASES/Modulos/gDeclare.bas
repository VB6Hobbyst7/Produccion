Attribute VB_Name = "gDeclare"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_FINDSTRING = &H14C
Global Const gsConnServDBF = "DSN=DSNCmactServ"
Global gsCentralPers As String

