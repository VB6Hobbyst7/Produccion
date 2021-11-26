Attribute VB_Name = "gManejadorError"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public Function GeneraMovNroPistas(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    'On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    Dim sql As String
    Set oConect = New DConecta
    Set rs = New ADODB.Recordset
    If oConect.AbreConexion = False Then Exit Function
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & Right(psCodAge, 2) & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = oConect.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNroPistas = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
'GeneraMovNroErr:
    'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Function GetMaquinaUsuario() As String
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    GetMaquinaUsuario = Trim(Left$(buffMaq, lSizeMaq))
End Function





