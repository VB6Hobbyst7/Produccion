VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum LogAreaTraEstado
    gLogAreaTraEstadoDesactivado = 0
    gLogAreaTraEstadoActivado = 1
    gLogAreaTraEstadoPrecio = 2
    gLogAreaTraEstadoCuenta = 3
    gLogAreaTraEstadoAcepta = 4
End Enum

Public Enum LogProvEstado
    gLogProvEstadoDesactivado = 0
    gLogProvEstadoActivado = 1
End Enum

Public Enum LogProvAgenteRetencion
    gLogProvAgenteRetencionNo = 0
    gLogProvAgenteRetencionSi = 1
End Enum

Public Enum LogProvBuenContribuyente
    gLogProvBuenContribuyenteNo = 0
    gLogProvBuenContribuyenteSi = 1
End Enum

Public Enum LogReqTipo
    gLogReqTipoNormal = 1
    gLogReqTipoExtemporaneo = 2
    gLogReqTipoConsolidado = 3
    gLogReqTipoSeleccion = 4
End Enum

Public Enum LogReqFlujo
    gLogReqFlujoSin = 1
End Enum

Public Enum LogReqEstado
    gLogReqEstadoInicio = 1
    gLogReqEstadoVB = 2
    gLogReqEstadoAcepPrevio = 3
    gLogReqEstadoPrecio = 4
    gLogReqEstadoCuenta = 5
    gLogReqEstadoConsolida = 6
    gLogReqEstadoParaTramite = 7
    gLogReqEstadoAceptado = 8
    gLogReqEstadoRechazado = 9
End Enum

'Public Enum LogObtEstado
'    gLogObtEstadoInicio = 1
'    gLogObtEstadoCuenta = 2
'    gLogObtEstadoAdquisi = 3
'    gLogObtEstadoAceptado = 8
'    gLogObtEstadoRechazado = 9
'End Enum

'Public Enum LogAdqEstado
'    gLogAdqEstadoInicio = 1
'    gLogAdqEstadoBase = 2
'    gLogAdqEstadoCotiza = 3
'    gLogAdqEstadoAceptado = 8
'    gLogAdqEstadoRechazado = 9
'End Enum

Public Enum LogSelEstado
    gLogSelEstadoInicioRes = 1
    gLogSelEstadoComite = 2
    gLogSelEstadoBases = 3
    gLogSelEstadoParametro = 4
    gLogSelEstadoPublicacion = 5
    gLogSelEstadoCotizacion = 6
    gLogSelEstadoRegBase = 7
    gLogSelEstadoAceptado = 8
    gLogSelEstadoRechazado = 9
    gLogSelEstadoDesierto = 10
    gLogSelEstadoProcAdju = 11
    gLogSelEstadoConsulta = 12
    gLogSelEstadoAbsolucion = 13
    gLogSelEstadoObservacion = 14
    gLogSelEstadoConsentimiento = 15
    gLogSelEstadoContratacion = 16
End Enum

Public Enum LogConEstado
    gLogConEstadoInicio = 1
    gLogConEstadoRecepcion = 2
    'gLogAdqEstadoAceptado = 8
    'gLogAdqEstadoRechazado = 9
End Enum

Public Enum LogSelSisAdj
    gLogSelSisAdjAdjDirecta = 1
    gLogSelSisAdjConPublico = 2
    gLogSelSisAdjLicPublica = 3
End Enum

Public Enum LogOperacion
    gLogOpeReqRegistro = 501001
    gLogOpeReqTramite = 501002
    gLogOpeReqModifica = 501003
    
    'gLogOpeObtRegistro = 501011
    'gLogOpeObtTramite = 501012
    'gLogOpeObtModifica = 501013
    
    'gLogOpeAdqRegistro = 501021
    'gLogOpeAdqTramite = 501022
    'gLogOpeAdqModifica = 501023
    
    gLogOpeSelRegistro = 501031
    gLogOpeSelTramite = 501032
    gLogOpeSelModifica = 501033
    
    gLogOpeConRegistro = 501041
    gLogOpeConTramite = 501042
    gLogOpeConModifica = 501043
    
    gLogOpeAlmRegistro = 501051
    gLogOpeAlmTramite = 501052
    gLogOpeAlmModifica = 501053
    
    'CONSTANTES DE ALMACEN REQUERIMIENTO
    gnAlmaReqAreaReg = 591001
    gnAlmaReqAreaMant = 591002
    gnAlmaReqAreaExt = 591003
    gnAlmaReqAreaRechPar = 591004
    
    gnAlmaIngXCompras = 591101
    gnAlmaIngXComprasConfirma = 591102
    gnAlmaIngXTransferencia = 591103
    gnAlmaIngXDevAreasGaran = 591104
    gnAlmaIngXProvGaranRepa = 591105
    gnAlmaIngXProvDemosOtros = 591106
    gnAlmaIngXDacionPago = 591107
    gnAlmaIngXEmbargo = 591108
    gnAlmaIngXAdjudicacion = 591109
    gnAlmaIngXOtrosMotivos = 591110
    
    gnAlmaSalXAtencion = 591201
    gnAlmaSalXTransferenciaOrigen = 591202
    gnAlmaSalXProvGarantRepa = 591203
    gnAlmaSalXAreasDevGaranRepa = 591204
    gnAlmaSalXProvDevCompras = 591205
    gnAlmaSalXProvDemosOtros = 591206
    gnAlmaSalXDevolEmbargo = 591207
    gnAlmaSalXOtrosMotivos = 591208
    gnAlmaSalXAjuste = 591209
    
    gnAlmaMantXIngreso = 591301
    gnAlmaMantXSalida = 591302
    
    gnAlmaExtornoXIngreso = 591401
    gnAlmaExtornoXSalida = 591402
    gnAlmaExtornoXConfirmacionIng = 591403
    
    gnAlmarReporteMovReq = 591501
    gnAlmarReporteMovNotIng = 591502
    gnAlmarReporteMovGuiaSal = 591503
    gnAlmarReporteListaIng = 591504
    gnAlmarReporteListaSal = 591505
    gnAlmarReporteResumenMovMes = 591506
    gnAlmarReporteIngAFBND = 591507
    gnAlmarReporteSalAFBND = 591508
    gnAlmarReporteActivoFijo = 591509
    
    gnAlmacenAsignaIGVNoUsado = 760499
End Enum


Public Enum LogOperacionCabecera
    gnLogOpeCab50 = 50
    gnLogOpeCab51 = 51
    gnLogOpeCab52 = 52
    gnLogOpeCab53 = 53
    gnLogOpeCab54 = 54
    gnLogOpeCab55 = 55
    gnLogOpeCab56 = 56
    gnLogOpeCab57 = 57
    gnLogOpeCab58 = 58
    gnLogOpeCab59 = 59
    gnLogOpeCabRequeri = 5910
    gnLogOpeCabIngresos = 5911
    gnLogOpeCabSalidas = 5912
    gnLogOpeCabModificaciones = 5913
    gnLogOpeCabExtornos = 5914
    gnLogOpeCabReportes = 5915
End Enum


Public Enum LogBSTpo
    gnLogBSTpoBien = 1
    gnLogBSTpoBienConsumo = 11
    gnLogBSTpoBienFijo = 12
    gnLogBSTpoBienNoDepreciable = 13
    gnLogBSTpoBienEmbargado = 14
    gnLogBSTpoBienAdjudicado = 15
    gnLogBSTpoBienDacionEnPago = 16
    gnLogBSTpoServicio = 2
End Enum


