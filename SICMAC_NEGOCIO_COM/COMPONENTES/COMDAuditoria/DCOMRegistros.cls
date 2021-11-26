VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub RegistrarActividad(ByVal cCodigoAct As String, ByVal cActividadDesc As String, ByVal nTipoAct As Integer, _
                                ByVal cFecRegistro As String, ByVal cNormaLegal As String, ByVal nOrigenAct As Integer)
    
    On Local Error Resume Next
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim pbTran As Boolean
    Dim lrDatos As ADODB.Recordset
    lsSQL = "exec stp_ins_RegistrarActividad '" & cCodigoAct & "','" & cActividadDesc & "'," & nTipoAct & ",'" & _
                                Format(cFecRegistro, "yyyy/MM/dd") & "','" & cNormaLegal & "'," & nOrigenAct & ""
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    If Err <> 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    loReg.CierraConexion
End Sub

Public Function ListarAuditTipoActividad() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_AuditTipoActividad"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ListarAuditTipoActividad = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Function ListarAuditOrigenActividad() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_AuditOrigenActividad"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ListarAuditOrigenActividad = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Function ObtenerAuditActividades() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_AuditDatosActividades"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ObtenerAuditActividades = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Function VerificarCodigoActividad(ByVal sCodigo As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_VerificaCodigoActividad '" & sCodigo & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set VerificarCodigoActividad = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Sub DarBajaActividad(ByVal sCodigo As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_upd_AuditDarBajaActividad '" & sCodigo & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub



Public Sub RegistrarAsignacionActividadUsuario(ByVal sCodigoAct As String, ByVal sPersCod As String, ByVal sObjetivoGeneral As String, _
                                               ByVal sObjetivoEspecifico As String, ByVal nEstado As Integer, ByVal sFechaRegistro As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_ins_RegistrarAsignacionActividadUsuario '" & sCodigoAct & "','" & sPersCod & "','" & sObjetivoGeneral & "','" _
                                                                 & sObjetivoEspecifico & "'," & nEstado & ",'" & sFechaRegistro & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ObtenerColaboradoresUAI() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ObtenerColaboradoresUAI"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ObtenerColaboradoresUAI = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Function AuditDatosActividadesXUser(ByVal sPersCod As String, ByVal nTipoUser As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_AuditDatosActividadesXUser '" & sPersCod & "'," & nTipoUser
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set AuditDatosActividadesXUser = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Sub AuditRegistrarProcedimientoActividad(ByVal cCodigoAct As String, ByVal cProcedimientoNombre As String, _
                                                ByVal cProcedimientoDesc As String, ByVal nEstado As Integer)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_ins_AuditRegistrarProcedimientoActividad '" & cCodigoAct & "','" & cProcedimientoNombre & "','" _
                                                                 & cProcedimientoDesc & "'," & nEstado
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ListarProcedimientosActividad(ByVal cCodigoAct As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ListarProcedimientosActividad '" & cCodigoAct & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ListarProcedimientosActividad = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Sub QuitarProcedimiento(ByVal nProcedimientoID As Integer)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_upd_QuitarProcedimiento " & nProcedimientoID
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Sub RegistrarValidacionProcedimiento(ByVal sValidacionCod As String, ByVal nProcedimientoID As String, ByVal sComentario As String, _
                                                 ByVal sConclusion As String, ByVal sFechaRegistro As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_ins_RegistrarValidacionProcedimiento '" & sValidacionCod & "'," & nProcedimientoID & ",'" & sComentario _
                                                              & "','" & sConclusion & "','" & Format(sFechaRegistro, "yyyy/MM/dd") & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function ObtenerCodigoValidacionProcedimiento() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
        
    lsSQL = "exec stp_sel_ObtenerCodigoValidacionProcedimiento"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ObtenerCodigoValidacionProcedimiento = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Function ListarProcedimientoVerificacion() As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ListarProcedimientoVerificacion"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ListarProcedimientoVerificacion = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Function ObtenerDatosValidacionProcedimiento(ByVal nProcedimientoID As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ObtenerDatosValidacionProcedimiento " & nProcedimientoID
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ObtenerDatosValidacionProcedimiento = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Sub ActualizarEstadoValidacionProcedimiento(ByVal nProcedimientoID As String, ByVal sValidacionCod As String, ByVal nTipoResolucion As Integer, ByVal sMotivoRechazo As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_upd_ActualizarEstadoValidacionProcedimiento " & nProcedimientoID & ",'" & sValidacionCod & "'," & nTipoResolucion & ",'" & sMotivoRechazo & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub

Public Function HistorialSolicitudValidacionProcedimento(ByVal nProcedimientoID As Integer) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_HistorialSolicitudValidacionProcedimento " & nProcedimientoID
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set HistorialSolicitudValidacionProcedimento = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Function ReporteSeguimientoActividades(ByVal sFecIni As String, ByVal sFecFin As String, ByVal sPersCod As String) As ADODB.Recordset
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    Dim lrDatos As ADODB.Recordset
    
    lsSQL = "exec stp_sel_ReporteSeguimientoActividades '" & Format(sFecIni, "yyyy/MM/dd") & "','" & Format(sFecFin, "yyyy/MM/dd") & "','" & sPersCod & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    Set lrDatos = loReg.CargaRecordSet(lsSQL, adLockReadOnly)
    Set ReporteSeguimientoActividades = lrDatos
    Set lrDatos = Nothing
    loReg.CierraConexion
End Function

Public Sub AuditCerrarActividad(ByVal cCodigoAct As String)
    Dim lsSQL As String
    Dim loReg As COMConecta.DCOMConecta
    
    lsSQL = "exec stp_upd_AuditCerrarActividad '" & cCodigoAct & "'"
    
    Set loReg = New COMConecta.DCOMConecta
    loReg.AbreConexion
    loReg.CargaRecordSet (lsSQL)
    loReg.CierraConexion
End Sub
