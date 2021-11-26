VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COMErrorHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



' This function will retrieve an error description from a resource
' file (.RES).  The ErrorNum is the index of the string
' in the resource file.  Called by RaiseError
Private Function GetErrorTextFromResource(ErrorNum As Long) As String
      On Error GoTo GetErrorTextFromResourceError
      Dim strMsg As String
     
      ' get the string from a resource file
      GetErrorTextFromResource = LoadResString(ErrorNum)
     
      Exit Function
GetErrorTextFromResourceError:
      If Err.Number <> 0 Then
            GetErrorTextFromResource = "An unknown error has occurred!"
      End If
End Function

'There are a number of methods for retrieving the error
'message.  The following method uses a resource file to
'retrieve strings indexed by the error number you are
'raising.
'---quitarcomentario
Public Sub RaiseError(ErrorNumber As Long, Source As String)
'      Dim strErrorText As String
'
'      'strErrorText = GetErrorTextFromResource(ErrorNumber)
'
'      'raise an error back to the client
'      'Err.Raise vbObjectError + ErrorNumber, Source, strErrorText
'      Err.Raise vbObjectError + ErrorNumber, Source, Err.Description
End Sub

Public Property Get MyUnhandledError() As Variant
    MyUnhandledError = 9999
End Property

