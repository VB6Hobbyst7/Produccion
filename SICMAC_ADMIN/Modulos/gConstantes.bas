Attribute VB_Name = "gConstantes"
Option Explicit

Global Const gsContDebe = "D"
Global Const gsContHaber = "H"
Global Const gsContDebeDesc = "Debe"
Global Const gsContHaberDesc = "Haber"

Global Const gsOpeCtaCaracterObligaDesc = "OBLIGATORIO"
Global Const gsOpeCtaCaracterOpcionDesc = "OPCIONAL"
Global Const gsOpeAnalCtaHisto = "82_02"

Global Const gsSI = "SI"
Global Const gsNO = "NO"

'Public BON   As String 'BOLD ON
'Public BOFF  As String 'Bold off
'Public CON   As String 'Condensado ON
'Public COFF  As String 'Condensado OFF

Public gcEntiOrig As String
Public gcEntiDest As String
'Public GSSIMBOLO As String
Public gcCtaEntiOrig As String
Public gcCtaEntiDest As String
Public gcDocNro As String


Public gcCentralCom As String
'ALPA 20090123***************************************
Public glsMovNro As String
'****************************************************
'Public gsOpeCod As String

'*******************************************************************
'*************** Seleccion de Personal

'************************** Anteriores (05-08)
'Global Const gsGeneracionCAP = "903032"
'Global Const gsGeneracionPAP = ""
'Global Const gsRegistroNivelesAprobacion = "903034"
'Global Const gsRegistroRequerimientoPersonal = "903035"
'Global Const gsRegistroAprobacionReqGerencia = "903036"
'Global Const gsRegistroAprobacionReqRecursos = "903037"
'Global Const gsMantenimientoFasesSeleccion = "903039"
'Global Const gsRegistroProcesoSeleccion = "903040"
'Global Const gsEnvioPropuestaGerencia = "903041"
'Global Const gsAprobacionPropuestaGerencia = "903042"
'Global Const gsInicioProcesoSeleccion = "903044"
'Global Const gsAprobacionPropuestaComite = "903043"
'Global Const gsRegistroFiltros = "903046"
'Global Const gsRegistroEvaluacionFases = "903047"
'Global Const gsRegistroAprobacionFasesComite = "903048"
'Global Const gsRegistroCostosPorFases = "903049"
'Global Const gsRegistroGanadoresSeleccion = "903052"
'Global Const gsCambiarMiembroTitular = "903051"
'Global Const gsAprobacionActaProceso = "903053"
'Global Const gsCambiaEstadoPostul = "903050"

Global Const gsGeneracionCAP = ""
Global Const gsGeneracionPAP = ""
Global Const gsRegistroRequerimientoPersonal = "903033"
Global Const gsRegistroNivelesAprobacion = "903032"
Global Const gsRegistroAprobacionReqGerencia = "903034"
Global Const gsMantenimientoFasesSeleccion = "903036"
Global Const gsRegistroProcesoSeleccion = "903037"
Global Const gsAprobacionPropuestaGerencia = "903038"
Global Const gsEstadoProcesoSeleccion = "903039"

Global Const gsCambiarMiembroTitular = "903040"

Global Const gsRegistroFiltros = "903042"
Global Const gsRegistroEvaluacionCurricular = "903043"
Global Const gsAprobacionEvaluacionCurricular = "903044"
Global Const gsRegistroEvaluacionConocimientos = "903045"
Global Const gsAprobacionEvaluacionConocimientos = "903046"
Global Const gsRegistroEvaluacionPsicologica = "903047"
Global Const gsAprobacionEvaluacionPsicologica = "903048"
Global Const gsRegistroEntrevistaPersonal = "903049"
Global Const gsAprobacionEntrevistaPersonal = "903050"
Global Const gsRegistroResultadosProceso = "903052"
Global Const gsCierreProcesoSeleccion = "903053"
Global Const gsImpresionActaYCuadro = "903054"
Global Const gsRegistroProcesoDesierto = "903056"
Global Const gsAprobacionProcesoDesierto = "903057"
Global Const gsActaProcesoDesierto = "903058"


'********** Contratacion de Personal************************
Global Const gsRegistroContratoSeleccion = "903061"
Global Const gsEntregaDocumentos = "903062"
Global Const gsRepEntregaDocumentos = "903063"
'Global Const gsRegistroFichaPersonal = "903065"
'Global Const gsReporteAperturaCuentas = "903064"
Global Const gsImpresionContratoPersonal = "903064"
Global Const gsRegistroFuncionesCargo = "903065"

'19-11-2005
Global Const gsRegistroRenovacionContrato = "903067"
Global Const gsRenovacionComentarioRRHH = "903068"
Global Const gsRenovacionVistoBuenoGerencia = "903069"

'************* Induccion de Personal *************************

Global Const gsRegistroPlantillaTemas = "903072"
Global Const gsRegistroTemas = "903071"
Global Const gsRegistroInduccion = "903073"
'Global Const gsRegistroAprobacionCronograma = "903074"
Global Const gsRegistroAsistenciaInduccion = "903074" '"903075"
Global Const gsRegistroEvaluacionAulas = "903075" '"903076"
Global Const gsRegistroEvaluacionCampo = "903076"
Global Const gsRegistroEvaluacionCampoFinal = "903077"
Global Const gsRegistroAsignacionFunciones = "903078" '"903077"

'************ Evaluacion de Desempeño ******************
Global Const gsRegistroCriteriosEvaluacion = "903081"
Global Const gsRegistroNivelesEvaluacion = "903082"
Global Const gsAperturaPeriodoEvaluacion = "903083"
Global Const gsRegistroEvaluacionPersonalPeriodica = "903084"
Global Const gsRegistroEvaluacionPersonalFinal = "903085"
Global Const gsComentarioRRHHEvaluacion = "903086"
Global Const gsVistoBuenoGerenciaEvaluacion = "903087"
Global Const gsHistoricoEvaluacionEmpleado = "903088"
Global Const gsConsultaEvaluacionFinal = "903089"

'************ Capacitacion de Personal *****************
Global Const gsRegistroCursosCapacitacion = "903091"
Global Const gsElaboracionPlanCapacitacion = "903092"
'Global Const gsRegNivelesAprobacionCapacitacion = "903093"
Global Const gsAprobacionPlanCapacitacion = "903093"
Global Const gsEjecucionPlanCapacitacion = "903094"

'*********************ccordova******************************'
Global Const gnEstadoProcSel = "9030"   'Estado de Proceso de Seleccion
Global Const gnTemaInduc = "9031"       'Tema de Induccion
Global Const gnEstadoInduc = "9032" 'Estado del Proceso de Induccion
Global Const gnTipoAsistSesi = "9033"   'Tipo se Asistencia a Sesiones
Global Const gnTipoCargaTema = "9034" 'Tipo de Carga Horaria en Temas
Global Const gnTipoInducTema = "9035"   'Tipo de Induccion en Temas
Global Const gnTipoAcepExpo = "9036"    'Tipo de Aceptacion de Expositor
Global Const gnTipoCargoComite = "9037"  'Tipo de Cargo de Comite
Global Const gnTipoRequisito = "9038"   'Tipo de Requisito
Global Const gnEstadoPostulante = "9039"   'Estado Postulante
Global Const gnEstadoRequerimientoPersonal = "9061"   'Estado Requerimiento Personal
Global Const gnRptOcupacionVivienda = "9062"
Global Const gnRptVivienda = "9063"
Global Const gnRptConfirmacion = "9064"

Global Const gnTipoClasificacion = "9071"
Global Const gnTipoEventoCurso = "9072"
Global Const gnTipoPublicoCurso = "9073"

Global Const gnNecesidadCapacitacion = "9074"
Global Const gnEstadoEvaluacionDesempeño = "9075"

Public Enum RHSeleccionFase
    RHSeleccionFaseCurricular = 6
    RHSeleccionFaseConocimientos = 8
    RHSeleccionFasePsicologica = 10
    RHSeleccionFaseEntrevista = 12
End Enum

Global Const gnModoPruebaRRHH = "140"
Global Const gnActualizarFichaPersonal = "141"

Public Enum LogPoderes
    LogPoderesNinguna = -1
    LogPoderesAprobacionADM = 1
    LogPoderesAprobacionJTer = 2
    LogPoderesAprobacionJLog = 3
    LogPoderesConfirmaADM = 4
    LogPoderesConfirmaJLog = 5
End Enum
'ALPA 20090126**************************************
Public Enum LogOperacionesPistas
    LogPistaMeritoDemerito = 557100
    LogPistaRegistraProcesoSeleccion = 550100
    LogPistaModificaProcesoSeleccion = 550200
    LogPistaConsultaProcesoSeleccion = 550300 'LUCV20181220, Anexo01 de Acta 199-2018
    LogPistaRegistraPostulante = 550500
    LogPistaModificaPostulante = 550600
    
    LogPistaModificaExamenCurricular = 551100
    LogPistaModificaExamenEscrito = 551500
    LogPistaModificaPsicologico = 551900
    LogPistaModificaExamenEntrevista = 552300
    
    LogPistaRegistraExamenCurricular = 551000
    LogPistaRegistraExamenEscrito = 551400
    LogPistaRegistraPsicologico = 551800
    LogPistaRegistraExamenEntrevista = 552200
    
    LogPistaRegistraCierreProcesoSelección = 552600
    
    LogPistaRegistraContratoProcesoSelección = 552700
    LogPistaModificaContratoProcesoSelección = 552900
    LogPistaRegistraContratoManual = 552800
    LogPistaModificaContratoManual = 552900
    
    LogPistaRescindirContrato = 553100
    
    LogPistaMantenimientoCurriculumTabla = 553300
    LogPistaRegistrarCurriculum = 553500
    LogPistaModificaCurriculum = 553600
    LogPistaIngresarSalirSistema = 700100
    
    'MAVM 20110407 ***
    LogPistaAdministracionUsuario = 550090
    '***
    
    'ARLO 20170126 **
    LogPistaIngresoSistema = 900104
    LogPistaConsultaPersona = 553000
    LogPistaEntraSalidaBienes = 501000
    LogPistaManProveedores = 558600
    LogPistaOrdenCompraSoles = 558700
    LogPistaOrdenCompraDolares = 558800
    LogPistaOrdenServicioSoles = 558900
    LogPistaOrdenServicioDolares = 559000
    LogPistaOrdenDolares = 558800
    LogPistaMantenientoOrdenCompraSoles = 501206
    LogPistaMantenientoOrdenCompraDolares = 502206
    LogPistaMantenientoOrdenServicioSoles = 501208
    LogPistaMantenientoOrdenServiciosDolares = 502208
    LogPistaseguientoOrdenes = 559110
    LogPistasImpresionOrdenComprasSoles = 501210
    LogPistasImpresionOrdenComprasDolares = 502210
    LogPistasImpresionOrdenServicioSoles = 501211
    LogPistasImpresionOrdenServicioDolares = 502211
    LogPistaRegistroContrato = 552800
    LogPistaRegistraAdenda = 553200
    LogPistaExtornoContrato = 552900
    LogPistaImprmirContrato = 552900
    LogPistaRegistroComprobanteMN = 591702
    LogPistaRegistroComprobanteME = 592702
    LogPistaComprobantes = 591700
    LogPistaExtornoComprobanteMN = 591703
    LogPistaExtornoComprobanteME = 592703
    LogPistaExtornoComprobanteLibreMN = 591704
    LogPistaExtornoComprobanteLibreME = 592704
    LogPistaActaConformidadMN = 591601
    LogPistaActaConformidadME = 592601
    LogPistaActaConformidadLibreMN = 591602
    LogPistaActaConformidadLibreME = 592602
    LogPistaInventarioAlmacen = 559100
    LogPistaKardexProducto = 559200
    LogPistaSaldosXAgencias = 158901
    LogPistaMantCuentaCont = 791600
    LogPistaMantemientoSaldos = 559400
    LogPistaReporteEstadistico = 780220
    LogPistaBajaActivo = 581299
    LogPistaEntradaSalidaBien = 502000
    LogPistaDepreActivoFijo = 581201
    LogPistaReportesActivoFijo = 763900
    '***
    
End Enum
'**********************************************************************
'WIOR 20121004*********************************************************
'CONSTANTES
Global Const gsLogContTipoPagoContratos = "10002"
Global Const gsLogContTipoContratos = "10028" 'Se cambio el valor 10003 a 10028 por PASI20140717 TI-ERS077-2014
Global Const gsLogContTipoGarantia = "10004"
Global Const gsLogContEstadoCuotas = "10043" 'Se cambio el valor 10005 a 10043 por PASI20140717 ERS0772014
Global Const gsLogContTipoAdendas = "10006"
Global Const gsLogContEstadoAdendas = "10007"
Global Const gsLogContEstadoContratos = "10008"
'RUTA DE CONTRATOS
Global Const gsLogContRutaContratos = "423"
'WIOR FIN *************************************************************
'EJVG20130520 ***
Public Enum LogTipoDocOrigenActivaBien
    OrdenCompra = 1
    OrdenServicio = 2
    CompraDirecta = 3
End Enum
Public Enum LogTipoBienesNew
    ActivoFijo = 1
    BienNoDepreciable = 2
    ActivoCompuesto = 3
    MejoraComponente = 4
    BienNoActivable = 5
End Enum
'END EJVG *******
'EJVG20131106 ***
'Global Const gsLogContTipoContratacion = "10028" 'Comentado PASI20140821 ERS0772014
Global Const gsLogContObjContrato = "10003" 'PASI20140717 ERS0772014
Global Const gsLogCtasContOCMN = "447"
Global Const gsLogCtasContOCME = "448"
Global Const gsLogCtasContOSMN = "449"
Global Const gsLogCtasContOSME = "450"
Public Enum LogTipoDocOrigenActaConformidad
    OrdenCompra = 1
    OrdenServicio = 2
    ContratoCompra = 3
    ContratoServicio = 4
    CompraLibre = 5
    Serviciolibre = 6
End Enum
'END EJVG *******
'PASI20140917 ERS0772014
Public Enum LogTipoDocOrigenComprobante
    OrdenCompra = 1
    OrdenServicio = 2
    ContratoServicio = 3
    ContratoArrendamiento = 4
    ContratoObra = 5
    ContratoAdqBienes = 6
    CompraLibre = 7
    Serviciolibre = 8
End Enum
Public Enum LogTipoContrato
    ContratoServicio = 3
    ContratoArrendamiento = 4
    ContratoObra = 5
    ContratoAdqBienes = 6
    ContratoSuministro = 7
End Enum
Public Enum LogtipoReajusteAdenda
    Complementaria = 1
    Adicional = 2
    Reduccion = 3
End Enum
'end PASI
