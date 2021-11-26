Attribute VB_Name = "gManejadorError"
Option Explicit
'Public objeto As New COMANEJADOR.ManejadorError

'Public Sub EnviarManejadorError(ByVal sProcedimiento As String, ByVal sSource As String)
'    Dim objBitacora As New COMNAuditoria.NCOMBitacora
'    Dim svar1 As String
'    Dim sNumMov As String
'    sNumMov = GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser)
'    objeto.ManejarError sProcedimiento, sSource, Err
'    objeto.Error Err
'    svar1 = objeto.PasarDatos
'    objeto.LimpiarError
'    MsgBox ("N° Error:" & vbCrLf & Err.Number & vbCrLf & vbCrLf & "Source:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Descripción:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Secuencia:" & vbCrLf & svar1), vbInformation, "Ha Ocurrido un Error"
'    objBitacora.InsertarPistaError "", sNumMov, gsCodPersUser, GetMaquinaUsuario, "4", Err.Number, Replace(Err.Source, "'", "") & ", " & Replace(Err.Description, "'", ""), svar1
'End Sub

Public Function GeneraMovNroPistas(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    'On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim sql As String
    Set oConect = New COMConecta.DCOMConecta
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
    'Exit Function
'GeneraMovNroErr:
    'Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

