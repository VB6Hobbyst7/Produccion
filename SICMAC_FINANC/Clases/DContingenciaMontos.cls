VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DContingenciaMontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CROB 20170726
Option Explicit
Dim dbConec As DConecta
Dim sSQL As String
Dim lsMsgErr As String
Dim rs As ADODB.Recordset
Enum TMontoRegistrado
    gMontoRegistrado = 1
    
End Enum

Public Sub RegistrarMontoContingencia(ByVal psNumRegistro As String, ByVal pnTipoMontoPCID As Integer, ByVal pnMonto As Double, ByVal pnMoneda As Integer _
                                    , ByVal pnTipoCambio As Double)
    On Error GoTo ERRORRegistrarMontoContingencia
    
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        sSQL = "exec stp_ins_RegistrarMontoContingencia '" & psNumRegistro & "'," & pnTipoMontoPCID & "," & pnMonto & "," & pnMoneda & "," & pnTipoCambio & ""
        dbConec.ConexionActiva.Execute sSQL
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Sub
    
ERRORRegistrarMontoContingencia:
    Err.Raise Err.Number, "Nuevo Monto Contingente", Err.Description
End Sub

'TORE08032018
Public Function NuevoMontoDetalle(ByVal psNumRegistro As String, ByVal pnTipoMontoPCID As Integer _
                                    , ByVal pnMonto As Double, ByVal pnMoneda As Integer _
                                    , ByVal pnTipoCambio As Double)
    On Error GoTo ErrorNuevoMontoDetalle
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "exec stp_ins_MontoRegistrarDetalle '" & psNumRegistro & "'," & pnTipoMontoPCID & "," & pnMonto & "," & pnMoneda & "," & pnTipoCambio & ""
    Set NuevoMontoDetalle = dbConec.CargaRecordSet(sSQL)
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Function
ErrorNuevoMontoDetalle:
        Err.Raise Err.Number, "Nuevo Monto Detalle", Err.Description
End Function


Public Function ActualizarMontoDetalle(ByVal psNumRegistro As String, ByVal pnTipoMontoPCID As Integer _
                                    , ByVal pnMonto As Double, ByVal pnMoneda As Integer _
                                    , ByVal pnTipoCambio As Double, ByVal pnTipoMontoPCID_Anterior As Integer)
On Error GoTo ErrorActualizarMontoDetalle
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "exec stp_upd_ActualizarMontoDetalleConsulta '" & psNumRegistro & "'," & pnTipoMontoPCID & "," & pnMonto & "," & pnMoneda & "," & pnTipoCambio & "," & pnTipoMontoPCID_Anterior & ""
    Set ActualizarMontoDetalle = dbConec.CargaRecordSet(sSQL)
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Function
ErrorActualizarMontoDetalle:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

Public Function EliminarMontoDetalle(ByVal psNumRegistro As String, ByVal pnTipoMontoPCID As Integer)
On Error GoTo ErrorEliminarMontoDetalle
    Set dbConec = New DConecta
    dbConec.AbreConexion
    sSQL = "exec stp_eli_EliminarConceptoCong '" & psNumRegistro & "'," & pnTipoMontoPCID & ""
    Set EliminarMontoDetalle = dbConec.CargaRecordSet(sSQL)
    dbConec.CierraConexion
    Set dbConec = Nothing
    Exit Function
ErrorEliminarMontoDetalle:
Call RaiseError(MyUnhandledError, "Error en proceso")
End Function

'END TORE
