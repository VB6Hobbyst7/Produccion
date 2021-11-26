VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'**********************Automatizacion del Formato de Revision
'**********************Validacion de la Calificacion

Public Function ValidarCierre(ByVal sFCierre As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ValidarCierre = objDAuditoria.DarValidarCierre(sFCierre)
End Function

Public Function InsertarRevision(ByVal sCodPers As String, ByVal sCodCta As String, ByVal sFCierre As String, ByVal CuTCambio As Currency, ByVal sFRegistro As String, ByVal sCAnalista As String, ByVal sGiro As String, ByVal sFSDCMAC As String, ByVal cMontoCMAC As Currency, ByVal sTMonedaCMAC As String, ByVal sFSDSF As String, ByVal cMontoSF As Currency, ByVal sTMonedaSF As String, ByVal sPNorm As String, ByVal sPCPP As String, ByVal sPDefic As String, ByVal sPDud As String, ByVal sPPerd As String, ByVal sCalificacionCMAC As String, ByVal sCalificacionSF As String, ByVal sCalificacionOCI As String, ByVal sSituacion As String, ByVal sDesarrollo As String, ByVal sGarantia As String, ByVal sInformacion As String, ByVal sEvaluacion As String, ByVal sComentario As String, ByVal sConclusion As String) As Integer
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRevision
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRevision
    objCOMDAuditoria.RegistrarRevisionCalificacion sCodPers, sCodCta, sFCierre, CuTCambio, sFRegistro, sCAnalista, sGiro, sFSDCMAC, cMontoCMAC, sTMonedaCMAC, sFSDSF, cMontoSF, sTMonedaSF, sPNorm, sPCPP, sPDefic, sPDud, sPPerd, sCalificacionCMAC, sCalificacionSF, sCalificacionOCI, sSituacion, sDesarrollo, sGarantia, sInformacion, sEvaluacion, sComentario, sConclusion
    InsertarRevision = objCOMDAuditoria.MaximoIdRevision
End Function

Public Sub EliminarRevisionCalificacion(ByVal iRevisionId As Integer)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRevision
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRevision
    objCOMDAuditoria.DeleteRevisionCalificacion iRevisionId
End Sub

'Public Function DarRevisionCalificacion(ByRef obj As COMANEJADOR.ManejadorError, ByVal psPersCod As String, ByVal pdFechaFinMes As Date, ByRef lsMensaje As String) As ADODB.Recordset ', ByVal pnTipCamb As Double
'On Error GoTo Manejador
'    Dim a, c As Integer
'    Dim objDAuditoria As COMDAuditoria.DCOMRevision
'    Set objDAuditoria = New COMDAuditoria.DCOMRevision
'    Set DarRevisionCalificacion = objDAuditoria.ListarRevisionCalificacion(obj, psPersCod, pdFechaFinMes, lsMensaje)
'    If obj.ValidarError Then GoTo Manejador
'    'a = c / 0
'    Exit Function
'
'Manejador:
'    obj.ManejarError "DarRevisionCalificacion", "COMNAuditoria/NCOMRevision", Err
'End Function

Public Function DarRevisionCalificacion(ByVal psPersCod As String, ByVal pdFechaFinMes As Date, ByRef lsMensaje As String) As ADODB.Recordset ', ByVal pnTipCamb As Double
'On Error GoTo Manejador
    'Dim a, c As Integer
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarRevisionCalificacion = objDAuditoria.ListarRevisionCalificacion(psPersCod, pdFechaFinMes, lsMensaje)
    'Dim obj As New COMANEJADOR.ManejadorError
    'If obj.ValidarError Then GoTo Manejador
    'a = c / 0
    'Exit Function

'Manejador:
    'obj.ManejarError "DarRevisionCalificacion", "COMNAuditoria/NCOMRevision", Err
End Function

Public Function DarCreditoXPersona(ByVal psPersCod As String, ByVal sFCierre As String, ByRef lmensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarCreditoXPersona = objDAuditoria.CargarCreditoXPersona(psPersCod, sFCierre, lmensaje)
End Function

Public Function ObtenerRevisionXId(ByVal iRevisionId As Integer) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerRevisionXId = objDAuditoria.DarRevisionXId(iRevisionId)
End Function

Public Function ObtenerDatosCalificacionXCliente(ByVal psPersCod As String, ByVal psCodCta As String, ByVal pdFechaFinMes As Date, ByVal pnTipCamb As Double, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosCalificacionXCliente = objDAuditoria.CargarDatosCalificacionXCliente(psPersCod, psCodCta, pdFechaFinMes, pnTipCamb, psMensaje)
End Function

'NR 20090815 MAVM
Public Function ObtenerPosicionClienteSBS(ByVal psPersId As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerPosicionClienteSBS = objDAuditoria.CargarPosicionClienteSBS(psPersId)
End Function

'NR 20090815 MAVM
Public Function ObtenerEndeudamientoSBS(ByVal psCodEnd As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerEndeudamientoSBS = objDAuditoria.CargarEndeudamientoSBS(psCodEnd)
End Function

Public Function ObtenerDatosCliente(ByVal psPersCod As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosCliente = objDAuditoria.DarDatosCliente(psPersCod)
End Function

Public Function ObtenerGarantia(ByVal psCodCta As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerGarantia = objDAuditoria.CargarGarantia(psCodCta, psMensaje)
End Function

Public Sub ActualizarRevisionCalificacion(ByVal iRevisionId As Integer, ByVal sCodPers As String, ByVal sCodCta As String, ByVal sFCierre As String, ByVal CuTCambio As Currency, ByVal sFRegistro As String, ByVal sCAnalista As String, ByVal sGiro As String, ByVal sFSDCMAC As String, ByVal cMontoCMAC As Currency, ByVal sTMonedaCMAC As String, ByVal sFSDSF As String, ByVal cMontoSF As Currency, ByVal sTMonedaSF As String, ByVal sPNorm As String, ByVal sPCPP As String, ByVal sPDefic As String, ByVal sPDud As String, ByVal sPPerd As String, ByVal sCalificacionCMAC As String, ByVal sCalificacionSF As String, ByVal sCalificacionOCI As String, ByVal sSituacion As String, ByVal sDesarrollo As String, ByVal sGarantia As String, ByVal sInformacion As String, ByVal sEvaluacion As String, ByVal sComentario As String, ByVal sConclusion As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRevision
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRevision
    objCOMDAuditoria.ModificarRevisionCalificacion iRevisionId, sCodPers, sCodCta, sFCierre, CuTCambio, sFRegistro, sCAnalista, sGiro, sFSDCMAC, cMontoCMAC, sTMonedaCMAC, sFSDSF, cMontoSF, sTMonedaSF, sPNorm, sPCPP, sPDefic, sPDud, sPPerd, sCalificacionCMAC, sCalificacionSF, sCalificacionOCI, sSituacion, sDesarrollo, sGarantia, sInformacion, sEvaluacion, sComentario, sConclusion
End Sub

Public Function ValRevision(ByVal psCodCta As String, ByVal pdFechaFinMes As String, ByRef lmensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ValRevision = objDAuditoria.ValidarRevision(psCodCta, pdFechaFinMes, lmensaje)
End Function

'Pistas

Public Function ObtenerDatosPistas(ByVal sFI As String, ByVal sFF As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosPistas = objDAuditoria.CargarDatosPistas(sFI, sFF, psMensaje)
End Function

Public Function ObtenerDatosPistasReporte(ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosPistasReporte = objDAuditoria.CargarDatosPistasReporte(sFI, sFF)
End Function

'******** MAVM Comentario de la Calificacion
Public Function DarAnalista(ByVal sAgencia As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarAnalista = objDAuditoria.CargarAnalista(sAgencia)
End Function

Public Function ObtenerInstituciones() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerInstituciones = objDAuditoria.CargarInstituciones()
End Function

Public Function ObtenerDatosCalificacionComentario(ByVal pdFechaFinMes As Date, ByVal pnTipCamb As Double, ByVal cCodInst As String, ByVal sFI As String, ByVal sFF As String, ByVal sImporteI As String, ByVal sImporteF As String, ByVal sAgencia As String, ByVal sAnalista As String, ByVal sMoneda As String, ByVal sCodTipoCred As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosCalificacionComentario = objDAuditoria.CargarDatosCalificacionComentario(pdFechaFinMes, pnTipCamb, cCodInst, sFI, sFF, sImporteI, sImporteF, sAgencia, sAnalista, sMoneda, sCodTipoCred)
End Function

' *** MAVM Generar Cartas de Circulacion ahorros
Public Function ObtenerDatosCtaAhorro(ByVal sNroCta As String, ByVal sPersCod As String, ByVal sAgencia As String, ByVal sCodProd As String, ByVal sMoneda As String, ByVal sImporteI As String, ByVal sImporteF As String, ByVal sTasa As String, ByVal sFI As String, ByVal sFF As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosCtaAhorro = objDAuditoria.CargarDatosCtaAhorro(sNroCta, sPersCod, sAgencia, sCodProd, sMoneda, sImporteI, sImporteF, sTasa, sFI, sFF, psMensaje)
End Function

Public Function ObtenerDatosXNroCta(ByVal sNroCta As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosXNroCta = objDAuditoria.DarDatosXNroCta(sNroCta)
End Function

Public Function ObtenerDatosCtaAhorroXCliente(ByVal psPersCod As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosCtaAhorroXCliente = objDAuditoria.CargarDatosCtaAhorroXCliente(psPersCod, psMensaje)
End Function

Public Function ObtenerProductos() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerProductos = objDAuditoria.CargarProductos
End Function

' *** MAVM Generar Cartas de Circulacion Credito
Public Function DarUserXCod(ByVal sCodPers As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarUserXCod = objDAuditoria.CargarUserXCod(sCodPers)
End Function

'NR Se Agrego los parametros moneda y monto de desmbolso
Public Function ObtenerDatosCredito(ByVal sNroCta As String, ByVal sCodPer As String, ByVal sAgencia As String, ByVal sCodTipCre As String, ByVal sAnalista As String, ByVal sFI As String, ByVal sFF As String, ByVal sMoneda As String, ByVal sImporteI As String, ByVal sImporteF As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosCredito = objDAuditoria.CargarDatosCredito(sNroCta, sCodPer, sAgencia, sCodTipCre, sAnalista, sFI, sFF, sMoneda, sImporteI, sImporteF, psMensaje)
End Function

Public Function ObtenerDatosFuncionario() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosFuncionario = objDAuditoria.DarDatosFuncionario()
End Function

Public Function ObtenerTipoCredito() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerTipoCredito = objDAuditoria.CargarTipoCredito
End Function

Public Function ObtenerDatosCreditoXCliente(ByVal psPersCod As String, ByRef psMensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerDatosCreditoXCliente = objDAuditoria.CargarDatosCreditoXCliente(psPersCod, psMensaje)
End Function

Public Function ObtenerCreditoXNroCta(ByVal sNroCta As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerCreditoXNroCta = objDAuditoria.DarCreditoXNroCta(sNroCta)
End Function

'*** Sistemas: Reporte de Usuarios y Perfiles

Public Function DarAgencias() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarAgencias = objDAuditoria.CargarAgencias()
End Function

Public Function DarAreas() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarAreas = objDAuditoria.CargarAreas()
End Function

Public Function DarDatosUsuarioXUser(ByVal psUser As String, ByVal psAgencia As String, ByVal psArea As String, ByRef lmensaje As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarDatosUsuarioXUser = objDAuditoria.CargarDatosUsuarioXUser(psUser, psAgencia, psArea, lmensaje)
End Function

Public Sub InsertarUsuarioCMACMAYNASTem(ByVal sNombre As String, ByVal sUsuario As String, ByVal sAgencia As String, ByVal sArea As String, ByVal sGrupo As String, ByVal sOperaciones As String, ByVal sColocaciones As String, ByVal sOtros As String)
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRevision
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRevision
    objCOMDAuditoria.RegistrarUsuarioCMACMAYNASTem sNombre, sUsuario, sAgencia, sArea, sGrupo, sOperaciones, sColocaciones, sOtros
End Sub

Public Sub BorrarUsuarioCMACMAYNASTem()
    Dim objCOMDAuditoria As COMDAuditoria.DCOMRevision
    Set objCOMDAuditoria = New COMDAuditoria.DCOMRevision
    objCOMDAuditoria.EliminarUsuarioCMACMAYNASTem
End Sub

Public Function DarUsuarioCMACMAYNASExcel() As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set DarUsuarioCMACMAYNASExcel = objDAuditoria.CargarUsuarioCMACMAYNASExcel
End Function

''R.A. Captaciones
'Public Function DarOperacionesXGrupo(ByVal psGrupo As String, ByVal svar1 As String, ByVal svar2 As String, ByRef lmensaje As String) As ADODB.Recordset
'    Dim objDAuditoria As COMDAuditoria.DCOMRevision
'    Set objDAuditoria = New COMDAuditoria.DCOMRevision
'    Set DarOperacionesXGrupo = objDAuditoria.CargarOperacionesXGrupo(psGrupo, svar1, svar2, lmensaje)
'End Function
'
''R.A. Colocaciones
'Public Function DarOperacionesXGrupoColocaciones(ByVal psGrupo As String, ByRef lmensaje As String) As ADODB.Recordset
'    Dim objDAuditoria As COMDAuditoria.DCOMRevision
'    Set objDAuditoria = New COMDAuditoria.DCOMRevision
'    Set DarOperacionesXGrupoColocaciones = objDAuditoria.CargarOperacionesXGrupoColocaciones(psGrupo, lmensaje)
'End Function
'
''R.A. Otros
'Public Function DarOperacionesXGrupoOtros(ByVal psGrupo As String, ByRef lmensaje As String) As ADODB.Recordset
'    Dim objDAuditoria As COMDAuditoria.DCOMRevision
'    Set objDAuditoria = New COMDAuditoria.DCOMRevision
'    Set DarOperacionesXGrupoOtros = objDAuditoria.CargarOperacionesXGrupoOtros(psGrupo, lmensaje)
'End Function

Public Function ObtenerCreditosDesembolsados(ByVal sFI As String, ByVal sFF As String) As ADODB.Recordset
    Dim objDAuditoria As COMDAuditoria.DCOMRevision
    Set objDAuditoria = New COMDAuditoria.DCOMRevision
    Set ObtenerCreditosDesembolsados = objDAuditoria.CargarCreditosDesembolsados(sFI, sFF)
End Function

'Public Function DarFechaCierreXAnioMes(ByVal sFecha As String) As ADODB.Recordset
'    Dim objDAuditoria As COMDAuditoria.DCOMRevision
'    Set objDAuditoria = New COMDAuditoria.DCOMRevision
'    Set DarFechaCierreXAnioMes = objDAuditoria.FechaCierreXAnio_Mes(sFecha)
'End Function








