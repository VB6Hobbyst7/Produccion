VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InsertarSesionDir(ByVal sTipoSesion, ByVal sNroSesion As String, ByVal sFecha As String) As Integer
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.RegistrarSesionDir sTipoSesion, sNroSesion, sFecha
    InsertarSesionDir = objCOMDAuditoria.MaximoIdSesionDir
End Function

Public Sub ActualizarSesionDir(ByVal iSesionDirId As Integer, ByVal sTipoSesion, ByVal sNroSesion As String, ByVal sFecha As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.ModificarSesionDir iSesionDirId, sTipoSesion, sNroSesion, sFecha
End Sub

Public Sub InsertarAcuerdoDir(ByVal iSesionDirId As Integer, ByVal sAsunto As String, ByVal sAcuerdo As String, ByVal sDetalle As String, ByVal sSituacion As String, ByVal tSituacion As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.RegistrarAcuerdoDir iSesionDirId, sAsunto, sAcuerdo, sDetalle, sSituacion, tSituacion
End Sub

Public Sub ActualizarAcuerdoDir(ByVal iSesionDirId As Integer, ByVal sNroAcuerdo As String, ByVal sAsunto As String, ByVal sAcuerdo As String, ByVal sDetalle As String, ByVal sSituacion As String, ByVal tSituacion As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.ModificarAcuerdoDir iSesionDirId, sNroAcuerdo, sAsunto, sAcuerdo, sDetalle, sSituacion, tSituacion
End Sub

Public Function BuscarAcuerdoDir(ByVal sDesde As String, ByVal sHasta As String, ByRef lsMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objDAuditoria = New COMDAuditoria.DCOMSeguimiento
    Set BuscarAcuerdoDir = objDAuditoria.ListarAcuerdoDir(sDesde, sHasta, lsMensaje)
End Function

Public Function ObtenerSesionDirXId(ByVal iSesionDirId As Integer, ByVal sNroAcuerdo As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objDAuditoria = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerSesionDirXId = objDAuditoria.DarSesionDirXId(iSesionDirId, sNroAcuerdo)
End Function

Public Function ObtenerReporteAcuerdoDir(ByVal sDesde As String, ByVal sHasta As String, ByVal sSituacion As String) As ADODB.Recordset
    Dim oInventario As COMDAuditoria.DCOMSeguimiento
    Set oInventario = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerReporteAcuerdoDir = oInventario.CargarReporteAcuerdoDir(sDesde, sHasta, sSituacion)
End Function

Public Function ObtenerReporteEstadistico() As ADODB.Recordset
    Dim oInventario As COMDAuditoria.DCOMSeguimiento
    Set oInventario = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerReporteEstadistico = oInventario.CargarReporteEstadistico
End Function


'Medidas de control

Public Function ObtenerTipoMedida() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objDAuditoria = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerTipoMedida = objDAuditoria.CargarTipoMedida
End Function

Public Function ObtenerEnte() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objDAuditoria = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerEnte = objDAuditoria.CargarEnte
End Function

Public Function InsertarInformeMed(ByVal sNroInforme As String, ByVal sFecha As String, ByVal sTipo As String) As Integer
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.RegistrarInformeMed sNroInforme, sFecha, sTipo
    InsertarInformeMed = objCOMDAuditoria.MaximoIdInformeMed
End Function

Public Sub ActualizarInformeMed(ByVal iInformeId As Integer, ByVal sTipoSesion, ByVal sNroSesion As String, ByVal sFecha As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.ModificarInformeMed iInformeId, sTipoSesion, sNroSesion, sFecha
End Sub

Public Sub InsertarMedidasCorrectivas(ByVal iInformeMedId As Integer, ByVal sObservacion As String, ByVal sRecomendacion As String, ByVal sAccion As String, ByVal sSituacion As String, ByVal tSituacion As String, ByVal sEnte As String, ByVal sArea As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.RegistrarMedidasCorrectivas iInformeMedId, sObservacion, sRecomendacion, sAccion, sSituacion, tSituacion, sEnte, sArea
End Sub

Public Sub ActualizarMedidasCorrectivas(ByVal iInformeId As Integer, ByVal iMedidasCorrectivasId As Integer, ByVal sObservacion As String, ByVal sRecomendacion As String, ByVal sAccion As String, ByVal sSituacion As String, ByVal tSituacion As String, ByVal sEnte As String, ByVal sArea As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objCOMDAuditoria = New COMDAuditoria.DCOMSeguimiento
    objCOMDAuditoria.ModificarMedidasCorrectivas iInformeId, iMedidasCorrectivasId, sObservacion, sRecomendacion, sAccion, sSituacion, tSituacion, sEnte, sArea
End Sub

Public Function BuscarMedidasCorrectivas(ByRef lsMensaje As String) As ADODB.Recordset 'ByVal sDesde As String, ByVal sHasta As String,
    Dim objDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objDAuditoria = New COMDAuditoria.DCOMSeguimiento
    Set BuscarMedidasCorrectivas = objDAuditoria.ListarMedidasCorrectivas(lsMensaje)
End Function

Public Function ObtenerInformeXId(ByVal iNroInformeId As Integer, ByVal iMedidasCorrectivasId As Integer) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMSeguimiento
    Set objDAuditoria = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerInformeXId = objDAuditoria.DarInformeXId(iNroInformeId, iMedidasCorrectivasId)
End Function

Public Function ObtenerReporteMedidasCorrectivas() As ADODB.Recordset 'ByVal sDesde As String, ByVal sHasta As String, ByVal sSituacion As String
    Dim oInventario As COMDAuditoria.DCOMSeguimiento
    Set oInventario = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerReporteMedidasCorrectivas = oInventario.CargarReporteMedidasCorrectivas() 'sDesde, sHasta, sSituacion
End Function

Public Function ObtenerReporteMedidasCorrectivasEstadistico() As ADODB.Recordset
    Dim oInventario As COMDAuditoria.DCOMSeguimiento
    Set oInventario = New COMDAuditoria.DCOMSeguimiento
    Set ObtenerReporteMedidasCorrectivasEstadistico = oInventario.CargarReporteMedidasCorrectivasEstadistico
End Function

