VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ManejadorError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub ManejarError(ByVal sProcedimiento As String, ByVal sSource As String, ByVal Error As ErrObject)
    Dim strErrDesc As String, strErrSource As String
    Dim strErrNumber As String, strExtraInfo As String
    strErrDesc = Error.Description
    strErrNumber = CStr(Error.Number)
    strErrSource = Error.Source
    If lsCadena = "" Then
        lsCadena = sSource & " " & sProcedimiento
    Else
        lsCadena = sSource & " " & sProcedimiento & vbCrLf & lsCadena
    End If
        If MatError(0, 1) = "" Then
            MatError(0, 0) = strErrNumber
            MatError(0, 1) = strErrSource
            MatError(0, 2) = strErrDesc
        End If
End Sub

Public Function PasarDatos() As String
    PasarDatos = lsCadena
End Function

Public Function ValidarError() As Boolean
    ValidarError = True
    If MatError(0, 0) <> "" Then
        ValidarError = True
    Else
       ValidarError = False
    End If
End Function

Public Sub Error(ByRef Error As ErrObject)
    Error.Number = IIf(MatError(0, 0) = "", 0, MatError(0, 0))
    Error.Source = MatError(0, 1)
    Error.Description = MatError(0, 2)
  
End Sub

Public Function LimpiarError()
    Erase MatError
    lsCadena = ""
End Function


