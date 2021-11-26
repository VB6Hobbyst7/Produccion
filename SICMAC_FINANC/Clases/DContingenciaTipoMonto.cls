VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DContingenciaTipoMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CROB 20170721
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim sSQL As String
Dim lsMsgErr As String
Dim rs As ADODB.Recordset

Public Sub RegistrarTipoMontoPasivoContingente(ByVal psDescripcion As String)
    On Error GoTo ERRORRegistrarTipoMontoPasivoContingente
    
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_ins_RegistrarTipoMontoPasivoContingente '" & psDescripcion & "'"
        dbConec.ConexionActiva.Execute sSQL
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
    
ERRORRegistrarTipoMontoPasivoContingente:
    Err.Raise Err.Number, "Nuevo Tipo Monto Pasivo Contingente", Err.Description
End Sub

Public Function ListarTipoMontoPasivoContingente()
    On Error GoTo ERRORListarTipoMontoPasivoContingente
    Set dbConec = New DConecta
    
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_sel_ListarTipoMontoPasivoContingente"
        Set ListarTipoMontoPasivoContingente = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
    
ERRORListarTipoMontoPasivoContingente:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Sub ActualizarTipoMontoPasivoContingente(ByVal pnCodigo As Integer, ByVal psDesc As String)
    On Error GoTo ERRORActualizarTipoMontoPasivoContingente
    
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_upd_ActualizarTipoMontoPasivoContingente " & pnCodigo & ", '" & psDesc & "'"
        dbConec.ConexionActiva.Execute sSQL
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
ERRORActualizarTipoMontoPasivoContingente:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

Public Function ListarTipoMoneda()
On Error GoTo ERRORListarTipoMoneda
    Set dbConec = New DConecta
    
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_sel_ListarTipoMoneda"
        Set ListarTipoMoneda = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ERRORListarTipoMoneda:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ObtenerTipoCambioFecha(ByVal psFecha As String)
On Error GoTo ERRORObtenerTipoCambioFecha
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_sel_RS_ObtieneTipoCambioFecha '" & Format(psFecha, "yyyymmdd") & "'"
        Set ObtenerTipoCambioFecha = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ERRORObtenerTipoCambioFecha:
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

'TORE08032017
Public Function NuevoMontosDetalle(ByVal pnNumeroRegistro As String, ByVal pnTipoMonto As Integer, _
                                         ByVal nMonto As Double, ByVal pnMoneda As Integer, ByVal nTipoCambio As Double)
On Error GoTo ErrorNuevoMontosDetalle
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_ins_RegistrarMontoContingencia '" & pnNumeroRegistro & "'," & pnTipoMonto & "," & nMonto & "," & pnMoneda & "," & nTipoCambio
        Set NuevoMontosDetalle = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorNuevoMontosDetalle:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function ActualizarMontosDetalles(ByVal pnNumeroRegistro As String, ByVal pnTipoMonto As Integer, _
                                         ByVal nMonto As Double, ByVal pnMoneda As Integer, ByVal nTipoCambio As Double, ByVal pnTipoMontoAnte As Integer)
On Error GoTo ErrorActualizarMontosDetalles
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_upd_ActualizarMontoDetalleConsulta '" & pnNumeroRegistro & "'," & pnTipoMonto & "," & nMonto & "," & pnMoneda & "," & nTipoCambio & "," & pnTipoMontoAnte
        Set ActualizarMontosDetalles = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorActualizarMontosDetalles:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function EliminarMontoDetalle(ByVal pnNumeroRegistro As String, ByVal pnTipoMonto As Integer)
On Error GoTo ErrorEliminarMontoDetalle
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_eli_EliminarMontoDetalleContingencia '" & pnNumeroRegistro & "'," & pnTipoMonto
        Set EliminarMontoDetalle = dbConec.CargaRecordSet(sSQL)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrorEliminarMontoDetalle:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

'END TORE


