VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub RegistrarActividad(ByVal cCodigoAct As String, ByVal cActividadDesc As String, ByVal nTipoAct As Integer, _
                                ByVal cFecRegistro As String, ByVal cNormaLegal As String, ByVal nOrigenAct As Integer)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.RegistrarActividad cCodigoAct, cActividadDesc, nTipoAct, cFecRegistro, cNormaLegal, nOrigenAct
End Sub

Public Function ListarAuditTipoActividad() As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ListarAuditTipoActividad = objCOMDAuditoria.ListarAuditTipoActividad
End Function

Public Function ListarAuditOrigenActividad() As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ListarAuditOrigenActividad = objCOMDAuditoria.ListarAuditOrigenActividad
End Function

Public Function ObtenerAuditActividades() As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ObtenerAuditActividades = objCOMDAuditoria.ObtenerAuditActividades
End Function

Public Function VerificarCodigoActividad(ByVal sCodigo As String) As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set VerificarCodigoActividad = objCOMDAuditoria.VerificarCodigoActividad(sCodigo)
End Function
Public Sub DarBajaActividad(ByVal sCodigo As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.DarBajaActividad (sCodigo)
End Sub

Public Sub RegistrarAsignacionActividadUsuario(ByVal sCodigoAct As String, ByVal sPersCod As String, ByVal sObjetivoGeneral As String, _
                                               ByVal sObjetivoEspecifico As String, ByVal nEstado As Integer, ByVal sFechaRegistro As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.RegistrarAsignacionActividadUsuario sCodigoAct, sPersCod, sObjetivoGeneral, sObjetivoEspecifico, nEstado, Format(sFechaRegistro, "yyyy/MM/dd")
End Sub

Public Function ObtenerColaboradoresUAI() As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ObtenerColaboradoresUAI = objCOMDAuditoria.ObtenerColaboradoresUAI
End Function

Public Function AuditDatosActividadesXUser(ByVal sPersCod As String, ByVal nTipoUser As Integer) As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set AuditDatosActividadesXUser = objCOMDAuditoria.AuditDatosActividadesXUser(sPersCod, nTipoUser)
End Function

Public Sub AuditRegistrarProcedimientoActividad(ByVal sCodigoAct As String, ByVal sProcedimientoNombre As String, _
                                                ByVal sProcedimientoDesc As String, ByVal nEstado As Integer)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.AuditRegistrarProcedimientoActividad sCodigoAct, sProcedimientoNombre, sProcedimientoDesc, nEstado
End Sub

Public Function ListarProcedimientosActividad(ByVal sCodigoAct As String) As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ListarProcedimientosActividad = objCOMDAuditoria.ListarProcedimientosActividad(sCodigoAct)
End Function

Public Sub QuitarProcedimiento(ByVal nProcedimientoID As Integer)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.QuitarProcedimiento (nProcedimientoID)
End Sub

Public Sub RegistrarValidacionProcedimiento(ByVal sValidacionCod As String, ByVal nProcedimientoID As String, ByVal sComentario As String, _
                                                 ByVal sConclusion As String, ByVal sFechaRegistro As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.RegistrarValidacionProcedimiento sValidacionCod, nProcedimientoID, sComentario, sConclusion, sFechaRegistro
End Sub

Public Function ListarProcedimientoVerificacion() As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ListarProcedimientoVerificacion = objCOMDAuditoria.ListarProcedimientoVerificacion
End Function

Public Function ObtenerDatosValidacionProcedimiento(ByVal nProcedimientoID As Integer) As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ObtenerDatosValidacionProcedimiento = objCOMDAuditoria.ObtenerDatosValidacionProcedimiento(nProcedimientoID)
End Function

Public Sub ActualizarEstadoValidacionProcedimiento(ByVal nProcedimientoID As String, ByVal sValidacionCod As String, ByVal nTipoResolucion As Integer, ByVal sMotivoRechazo As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.ActualizarEstadoValidacionProcedimiento nProcedimientoID, sValidacionCod, nTipoResolucion, sMotivoRechazo
End Sub

Public Function HistorialSolicitudValidacionProcedimento(ByVal nProcedimientoID As Integer) As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set HistorialSolicitudValidacionProcedimento = objCOMDAuditoria.HistorialSolicitudValidacionProcedimento(nProcedimientoID)
End Function

Public Function ReporteSeguimientoActividades(ByVal sFecIni As String, ByVal sFecFin As String, ByVal sPersCod As String) As ADODB.Recordset
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Set ReporteSeguimientoActividades = objCOMDAuditoria.ReporteSeguimientoActividades(sFecIni, sFecFin, sPersCod)
End Function

Public Sub AuditCerrarActividad(ByVal cCodigoAct As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    objCOMDAuditoria.AuditCerrarActividad (cCodigoAct)
End Sub

Public Function ObtenerCodigoValidacionProcedimiento() As String
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRegistros
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRegistros
    Dim lrDatos As ADODB.Recordset
    Set lrDatos = New ADODB.Recordset
    
    Set lrDatos = objCOMDAuditoria.ObtenerCodigoValidacionProcedimiento()
    Do Until lrDatos.EOF
        ObtenerCodigoValidacionProcedimiento = lrDatos!cValidacionCod
        lrDatos.MoveNext
    Loop
    If ObtenerCodigoValidacionProcedimiento = "" Then
        ObtenerCodigoValidacionProcedimiento = "AUD-10001"
    Else
        Dim sTmpTexto As String
        sTmpTexto = Mid(ObtenerCodigoValidacionProcedimiento, 5, Len(ObtenerCodigoValidacionProcedimiento) - 1)
        ObtenerCodigoValidacionProcedimiento = "AUD-" & (sTmpTexto + 1)
    End If
End Function
