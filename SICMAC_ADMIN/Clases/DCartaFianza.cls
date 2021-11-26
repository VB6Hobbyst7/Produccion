VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCartaFianza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim oConecta As DConecta
Dim rs As ADODB.Recordset
Dim sSql As String
 

Public Function RegistraCartaFianza(ByVal psNroCartaFianza As String, ByVal psCodProv As String, ByVal psNroContrato As String, _
                                    ByVal psDescripcion As String, ByVal pCodIFI As String, ByVal pnMoneda As Integer, _
                                    ByVal pdFEmision As String, ByVal psFVencimiento As String, ByVal psUltimaActulizacion As String, _
                                    Optional ByVal pnEstado As Integer = 0)
    On Error GoTo ErrorRegistraCartaFianza
    
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "exec stp_ins_upd_RegistrarCartaFianza '" & psNroCartaFianza & "','" & psCodProv & "','" & psNroContrato & "','" & psDescripcion & "','" _
                                                & pCodIFI & "'," & pnMoneda & ",'" & pdFEmision & "','" & psFVencimiento & "','" & psUltimaActulizacion & "'," & pnEstado
    Set RegistraCartaFianza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorRegistraCartaFianza:
    Err.Raise Err.Number, "Registro de Carta Fianza", Err.Description
End Function

Public Function ObtenerInfoCartaFianzaNroCF(ByVal psNroCartaFianza As String) As ADODB.Recordset
    On Error GoTo ErrorObtenerInfoCartaFianza
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "exec stp_sel_ObtenerCartaFianzaNroCF '" & psNroCartaFianza & "'"
    Set ObtenerInfoCartaFianzaNroCF = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInfoCartaFianza:
    Err.Raise Err.Number, "Obtener Info Carta Fianza por Nro Carta", Err.Description
End Function

Public Function ObtenerInfoCartaFianzaCodProv(ByVal psCodProveedor As String, _
                                                ByVal pdFDesde As String, _
                                                ByVal pdFHasta As String) As ADODB.Recordset
    On Error GoTo ErrorObtenerInfoCartaFianzaCodProv
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "exec stp_sel_ObtenerCartaFianzaCodProveedor '" & psCodProveedor & "','" & pdFDesde & "','" & pdFHasta & "'"
    Set ObtenerInfoCartaFianzaCodProv = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerInfoCartaFianzaCodProv:
    Err.Raise Err.Number, "Obtener Info Carta Fianza por Cod Proveedor", Err.Description
End Function

Public Function ObtenerValidacionCartaFianza(ByVal psNroCF As String, ByVal psNroContrato As String, ByVal pnAccionar As Integer) As String
    On Error GoTo ErrorObtenerValidacionCartaFianza
    Set oConecta = New DConecta
    Set rs = New ADODB.Recordset
    
    oConecta.AbreConexion
    sSql = "exec stl_sel_ObtenerValidacionCartaFianza '" & psNroCF & "','" & psNroContrato & "'," & pnAccionar
    Set rs = oConecta.CargaRecordSet(sSql)
    If Not rs.BOF And Not rs.EOF Then
        ObtenerValidacionCartaFianza = rs!cMsjValidacion
    Else
        ObtenerValidacionCartaFianza = ""
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorObtenerValidacionCartaFianza:
    Err.Raise Err.Number, "Obtener Info Carta Fianza por cod prov", Err.Description
End Function

Public Function EnvioCorreoCartaFianza(ByVal psUser As String, ByVal pdFechaSist As String) As ADODB.Recordset
    On Error GoTo ErrorEnvioCorreoCartaFianza
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = "exec stp_sel_VerificarEnvioCorreoCartaFianza '" & psUser & "','" & pdFechaSist & "'"
    Set EnvioCorreoCartaFianza = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    Exit Function
ErrorEnvioCorreoCartaFianza:
    Err.Raise Err.Number, "Envio correo de carta fianzas vencidas", Err.Description
End Function
