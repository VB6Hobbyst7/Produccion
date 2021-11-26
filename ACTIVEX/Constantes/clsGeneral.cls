VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ConstanteCabecera
    gProducto = 1001
    gPersPersoneria = 1002
    gPersIdTipo = 1003
    gPersJurMagnitud = 1004 '= "gsPersJurMagnitud"
    gPersEstado = 1005 '= "gsPersEstado"
    gPersRelacion = 1006 '= gsPersRelaC"
    gChequeEstado = 1007
    gProductoCuentaTipo = 1008
    gChequePlaza = 1009
    gMeses = 1010
    gMoneda = 1011
    gTipoCambio = 1012
    gTipoCalculoInteres = 1013
    gProductoParametro = 1014
    gMotivoNotaAbonoCargo = 1015
    gUnidadMedida = 1019
    gPersEstadoCivil = 1020
    gPersRelGarantia = 1021
    gDiasSemana = 1022
    gPersRelacBenef = 1023
    gPersRelacAmp = 1024
    gPersRelacEstado = 1025
    gPersTipo = 1026
    gPersGarantia = 1027
    gPersFteIngresoTipo = 1028
    gPersCondDomic = 1029
    gPersGarantEstado = 1030
    gGruposIngEgre = 1031
    gNotaCargoAbonoEstado = 1032
    gTipoDiaHabil = 1033
    gTipoCierreRealizado = 1034
    gPersRelacionInst = 1035 'Relacion con la Institucion
    gTipoPeriferico = 1036
    gTipoPinPad = 9050
    
    gTipoPuertoSerial = 1037
    gPersEstLavDinero = 1038
    gPersTpoSangre = 1039
    gGarantTpoInmueb = 1040
    gGastosProd = 1041
    gChequeTipo = 1042
    gAreasAgrupacion = 1043  'Agrupacion de Areas
    gTipoExoneracionITF = 1044  'Agrupacion de Areas
    
    gCaptacEstado = 2001
    gCaptacTipoTasa = 2002
    gCaptacPFFormaRetiro = 2003
    gCaptacOrdPagoEstado = 2004
    gCaptacRelacPersona = 2005
    gCaptacTarjetaEstado = 2006
    gCaptacMotBloqueoRet = 2007
    gCaptacMotBloqueoTot = 2008
    gCaptacTipoBloqueo = 2009
    gCaptacParametro = 2010
    gCaptacConcepto = 2011
    gCaptacMovTipo = 2012
    gCaptacInstServicios = 2013
    gCaptacFlagServicios = 2014
    gCaptacConvenios = 2015
    gCaptacConvTipoCuenta = 2016
    gCaptacConvTipoReferencia = 2017
    gCapServUNTConcepto = 2018
    gCapTpoOperacionEstad = 2019
    gCapPermisosEspeciales = 2020
    gCapOrdPagTalEstado = 2021
    gCapOrdPagTalTipo = 2022
    gCaptacTipoEstado = 2023
    gCapServTipoComision = 2024
    
    gCapEstAutorizaOpe = 9051  'ppoa
    
    
    gServConcepto = 2025
    gCapTasaEspEstados = 2026
    gColocEstado = 3001
    gColocRelacPers = 3002
    gColocTipoTasa = 3003
    gColocGrupo = 3004
    gColocTipoCalend = 3005
    gColocCalendApl = 3006
    gColocCalendEstado = 3007
    gColocConcepto = 3008
    gColocConceptoApl = 3009
    gColocCalific = 3010
    gColocNota = 3011
    gColocLineaCredTasas = 3012
    gColocLineaCredPlazo = 3013
    gColocCredTipo = 3014
    gColocCredCondicion = 3015
    gColocCredCondicionOtra = 3030
    gColocDestino = 3016
    gColocTiposGracia = 3017
    gColocTiposDesembolso = 3018
    gColocCalendConceptoProc = 3019
    gColocTipoPago = 3020
    gColocEstadoRefinanc = 3021
    gColocMetasCred = 3022
    gColocTipoMetas = 3023
    gColocMotivRechazo = 3024
    gColocGarantRec = 3025
    gColocGarantRecEstado = 3026
    'gColCFModalidad = 3027
    gColocMiVivTramo = 3028
    gColocParam = 3029
    
    'Pigno
    gColocPParametros = 3201
    gColocPRecGarEstado = 3202
    gColocPRecGarVentaEstado = 3203
    gColocPOperaciones = 3204
    gColocPReportes = 3205
    gColocPMaterialJoyas = 3206
    'Recup
    gColocRecTipoCobranza = 3301
    gColocRecDemandado = 3302
    gColocRecViaProcesal = 3303
    gColocRecTpoComision = 3304
    gColocRecGastoEstado = 3305
    gColocRecExpedTipoInf = 3306
    gColocRecOperaciones = 3307
    gColocRecExpedienteEstado = 3309
    'CF
    gColCFModalidad = 3402
    gColocCFOperaciones = 3403
    gColocCFReportes = 3403
    gColocCFMotivHonrar = 3404
    
    'Garantias
    gGarProcesos = 3501
    
    
    gCGTipoBloqueos = 4000
    gCGTipoIF = 4001
    gCGTipoCtaIF = 4002
    gCGTipoEstCuotaAdeud = 4003
    gCGTipoPagoAdeud = 4004
    gCGTipoCuotCalIF = 4005
    gMovEstado = 4006
    gMovFlag = 4007  ' "gsMovFlag"
    gMovParalelo = 4008 '= "CMovParalelo"
    gCGEstadoCtaIF = 4009
    gCHTipoProc = 4010
    gArendirTipo = 4011
    gViaticosCateg = 4012
    gViaticosDestino = 4013
    gViaticosAfectoA = 4014
    gViaticosTransporte = 4015
    gOpeTpoVisible = 4016
    gOpeCtaCaracter = 4017
    gFiltroObj = 4018
    gCGEstadosOp = 4019
    gFSDTipoExonerados = 4020
    gTpoObjetos = 4021
    gCGEstadoConfCheque = 4022
    gCGTipoTransporte = 4023
    gCGCtaIFConceptos = 4024
    gAnxTipoCalculo = 4025
    gCGAdeudCalTpoCuota = 4030
    gAnx10TipoOficina = 4031
    gViaticosMotivos = 4032
    
    gLogProvEstado = 5001
    gLogReqTipo = 5002
    gLogReqEstado = 5003
    gLogReqFlujo = 5004
    gLogOperacion = 5005
    'gLogObtEstado = 5006
    'gLogAdqEstado = 5007
    gLogSelEstado = 5008
    gLogSelTpo = 5051
    gLogSelParTec = 5052
    gLogSelParEco = 5053
    gLogSelCro = 5054
    gLogSelSisAdj = 5055
    gLogSelParTpo = 5056
    gLogSerConTpo = 5057
    
    
    gLogProvAgeReten = 5060
    gLogProvBuenCont = 5061
    
    
    gRHBeneficiarioRela = 6001
    gRHAsistenciaMedicaRela = 6002
    gRHProcesoSeleccionTipo = 6003
    gRHProcesoSeleccionEstado = 6004
    gRHProcesoSeleccionResultado = 6005
    gRHEstado = 6006
    gRHCondicion = 6007
    gRHCategoria = 6008
    gRHCargoEstado = 6009
    gRHConceptoIngresos = 6010
    gRHConceptoDescuentos = 6011
    gRHTipoContrato = 6012
    gRHConceptoVariablesUsuario = 6013
    gRHConceptoVariablesGlobales = 6014
    gRHConceptoFuncionesConstantes = 6015
    gRHConceptoVariablesLocales = 6016
    gRHConceptoTipo = 6017
    gRHConceptoEstado = 6018
    gRHImpuesto5taCat = 6019
    gRHAfeCtoMesTrabajado = 6020
    gRHProcesoSeleccionModal = 6021
    gRHPermisosCab = 6025
    gRHPlanillaTipoCalculo = 6028
    gRHDemeritosCab = 6035
    gGiroRelacPersona = 9001
    gGiroTarifaTipo = 9002
    
    'Pignoraticio Lima
    gColocPigTipoJoya = 8001
    gColocPigSubTipoJoya = 8002
    gColocPigMaterial = 8003
    gColocPigUbicacion = 8004
    gColocPigCalifCte = 8005
    gColocPigEstConservaJoya = 8006
    gColocPigTipoProcesoTar = 8008
    gColocPigTipoProceso = 8009
    gColocPigTipoGuia = 8010
    gColocPigTipoDocumento = 8011
    gColocPigSituacionPieza = 8012
    gColocPigTipoTasacion = 8013
    gColocPigCalifSbs = 8014
    gColocPigMotivoRem = 8015
    gColocPigTipoVentaJoya = 8016
    
    gColocPigPlazoMin = 8015        'Parametros
    gColocPigPlazoMax = 8016        'Parametros
    
End Enum

Public Enum Producto
    gCapTodoAhorro = 99
    gCapAhorros = 232
    gCapPlazoFijo = 233
    gCapCTS = 234
    gGiro = 239
    gColComercEmp = 101
    gColComercAgro = 102
    gColComercPesq = 103
    gColPYMEEmp = 201
    gColPYMEAgro = 202
    gColPYMEPesq = 203
    gColConsuDctoPlan = 301
    gColConsuPlazoFijo = 302
    gColConsCTS = 303
    gColConsuUsosDiv = 304
    gColConsuPrendario = 305
    gColConsuPrestAdm = 320
    gColCFComercial = 121
    gColCFPYME = 221
    gColHipoCaja = 401
    gColHipoMiVivienda = 423
End Enum

Enum GastosProd
    gGastosProdCredito = 1
    gGastosProdJudicial = 2
End Enum

Public Enum PersTpoSangre
    gPersTpoSangreNODEF = 0
    gPersTpoSangreARHP = 1
    gPersTpoSangreARHN = 2
    gPersTpoSangreBRHP = 3
    gPersTpoSangreBRHN = 4
    gPersTpoSangreABRHP = 5
    gPersTpoSangreABRHN = 6
    gPersTpoSangreORHP = 7
    gPersTpoSangreORHN = 8
End Enum

Public Enum GarantTpoInmueb
    gGarantTpoInmuebCasaInd = 1
    gGarantTpoInmuebCasaCond = 2
    gGarantTpoInmuebCasaDpto = 3
    gGarantTpoInmuebProyInmob = 4
End Enum

Public Enum PersPersoneria  '= "gsPersPersoneria"
    gPersonaNat = 1
    gPersonaJurSFL = 2
    gPersonaJurCFL = 3
    gPersonaJurCFLCMAC = 4
    gPersonaJurCFLCRAC = 5
    gPersonaJurCFLFONCODES = 6
    gPersonaJurCFLCooperativa = 7
    gPersonaJurCFLEdpyme = 8
    
End Enum

Public Enum PersIdTipo  '= "gsPersIdTpo"
    gPersIdDNI = 1  ' "DNI"
    gPersIdRUC = 2  '"RUC"
    gPersIdFFPPAA = 3 ' "FPA"
    gPersIdExtranjeria = 4 ' "EXT"
    gPersIdSBS = 5 ' "SBS"
    gPersIdIPSS = 6 ' "IPS"
    gPersIdBrevete = 7 ' "BRE"
    gPersIdRegPub = 8   '"RPU"
    gPersIdAFP = 9    ' "AFP"
    gPersIdPartNaC = 10 ' "PAR"
    gPersIdPasaporte = 11 ' "PAS"
    gPersIdBolMil = 12  '"BOL"
    gPersIdLibMil = 13 ' "LIB"
    gPersIdLibTribut = 14 ' "LTB"
    gPersIdRUS = 15  '"RUS"
End Enum

Public Enum PersJurMagnitud  '= "gsPersJurMagnitud"
    gPersJurMagnitudGrande = 0
    gPersJurMagnitudMediana = 1
    gPersJurMagnitudPequeña = 2
    gPersJurMagnitudMicro = 3
End Enum

Public Enum PersEstado '= "gsPersEstado"
    gPersPNActivo = 0
    gPersPNFallecido = 1
    gPersPJActivo = 10
    gPersPJLiquidacion = 11
    gPersPJDisuelta = 12
    gPersPJFusionada = 13
End Enum

Public Enum PersRelacion  '= gsPersRelaC"
    gPersRelConyugue = 0
    gPersRelHijo = 1
    gPersRelPadre = 2
    gPersRelHermano = 3
    gPersRelOtro = 4
End Enum

Enum PersRelacEstado
    gPersRelacEstadoActivo = 1
    gPersRelacEstadoNOActivo = 0
End Enum

Enum PersEstadoCivil
    gPersEstadoCivilSoltero = 1
    gPersEstadoCivilCasado = 2
    gPersEstadoCivilViudo = 3
    gPersEstadoCivilDivorciado = 4
    gPersEstadoCivilSeparado = 5
    gPersEstadoCivilConviviente = 6
    gPersEstadoCivilNOIndicado = 0
End Enum

Enum PersCondDomic
    gPersCondDomicPropio = 1
    gPersCondDomicAlquilado = 2
    gPersCondDomicEnUso = 3
    gPersCondDomicNoIndicado = 0
End Enum

Enum PersFteIngresoTipo
    gPersFteIngresoTipoDependiente = 1
    gPersFteIngresoTipoIndependiente = 2
End Enum

Enum PersRelacBenef
    gPersRelacBenefActivo = 1
    gPersRelacBenefNOActivo = 0
End Enum

Enum PersRelacAMP
    gPersRelacAMPActiva = 1
    gPersRelacAMPNOActiva = 0
End Enum

Enum PersTipo
    gPersTipoConvenio = 1
End Enum

Enum PersGarantia
    gPersGarantiaHipotecas = 19
    gPersGarantiaPrendaVehicular = 2
    gPersGarantiaPrendaIndustrial = 3
    gPersGarantiaPrendaAgricola = 4
    gPersGarantiaCartasFianza = 5
    gPersGarantiaDepositosGarantia = 6
    gPersGarantiaMetalesPreciososJoyas = 7
    gPersGarantiaPrendaGlobal = 8
    gPersGarantiaFianzaSolidaria = 9
    gPersGarantiaOtrasGarantias = 10
    gPersGarantiaPatrimonioPersonal = 11
    gPersGarantiaCartaAutorizacionDescuento = 12
End Enum

Enum PersGarantEstado
    gPersGarantEstadoRegistrado = 1
    gPersGarantEstadoAsignado = 2
    gPersGarantEstadoContabilizado = 3
    gPersGarantEstadoLiberado = 4
    gPersGarantEstadoLevantada = 5
    gPersGarantEstadoBloqueada = 6
End Enum

Enum PersRelGarantia
    gPersRelGarantiaTitular = 1
    gPersRelGarantiaGarante = 2
    gPersRelGarantiaConyugue = 3
    gPersRelGarantiaRepresentante = 4
End Enum

Enum PersRelacionInst
    gPersRelacionInstDirector = 1
    gPersRelacionInstFuncionario = 2
    gPersRelacionInstNinguna = 3
    gPersRelacionInstNOIndicado = 4
    gPersRelacionInstTrabajador = 5
End Enum

Public Enum ChequeEstado
    gChqEstRegistrado = 0
    gChqEstEnValorizacion = 1  'E
    gChqEstValorizado = 2  'V
    gChqEstAnulado = 3 'A
    gChqEstRechazado = 4   'R
    gsChqEstExtornado = 5   'X
End Enum

Public Enum NotaCargoAbonoEstado
    gNCNARegistrado = 0
    gNCNAEnMovimiento = 1
    gNCNAPorConfirmar = 2
    gNCNAConfirmado = 3
    gNCNAAnulado = 4
    gNCNARechazado = 5
End Enum

Public Enum ChequeTipo
    gChqTpoSimple = 0
    gChqTpoGerencia = 1
End Enum

Public Enum ChequePlaza
    gChqPlazaLocal = 0
    gChqPlazaOtros = 1
End Enum

Public Enum Moneda
    gMonedaNacional = 1
    gMonedaExtranjera = 2
    gMonedaVAC = 3
End Enum

Public Enum ProductoCuentaTipo
    gPrdCtaTpoIndiv = 0 'I
    gPrdCtaTpoMancom = 1 'Y
    gPrdCtaTpoIndist = 2 'O
End Enum

Public Enum ProductoParametro
    gPrdParamColoc = 1000
    gPrdParamCaptac = 2000
    gPrdParamContab = 4000
    gPrdParamLogist = 5000
    gPrdParamRRHH = 6000
End Enum

Public Enum Meses
    gMesEnero = 1
    gMesFebrero = 2
    gMesMarzo = 3
    gMesAbril = 4
    gMesMayo = 5
    gMesJunio = 6
    gMesJulio = 7
    gMesAgosto = 8
    gMesSeptiembre = 9
    gMesOctubre = 10
    gMesNoviembre = 11
    gMesDiciembre = 12
End Enum

Public Enum UnidadMedida
    gUniMedUnidad = 1
    gUniMedKilo = 2
    gUniMedMetro = 3
    gUniMedLitro = 4
End Enum

Public Enum TipoCambio
    TCFijoMes = 0
    TCFijoDia = 1
    TCVenta = 2
    TCCompra = 3
    TCVentaEsp = 4
    TCCompraEsp = 5
    TCPonderado = 6
End Enum

Public Enum GruposIngEgre
    gGruposIngEgreAgeLocal = 0
    gGruposIngEgreOtraAgencia = 1
    gGruposIngEgreOtraCMAC = 2
    gGruposIngEgreCompraVenta = 3
    gGruposIngEgreServicios = 4
    gGruposIngEgreOtrasOpe = 5
    gGruposIngEgreHabDev = 6
    gGruposIngEgreSobFalt = 7
End Enum


Public Enum TpoDoc
    TpoDocFactura = 1
    TpoDocRecHonorarios = 2
    TpoDocBoletaVenta = 3
    TpoDocLiquidCompra = 4
    TpoDocBolCiasAciacion = 5
    TpoDocCartaPorteAereo = 6
    TpoDocNotaCredito = 7
    TpoDocNotaDebito = 8
    TpoDocBolViajeTerrestre = 9
    TpoDocRecArriendo = 10
    TpoDocPolizaBolsaValores = 11
    TpoDocTicketMaqReg = 12
    TpoDocDocBancaySeguro = 13
    TpoDocRecServPublicos = 14
    TpoDocDocAFPSyEPS = 18
    TpoDocDocInstPublicas = 19
    TpoDocGuiaRemision = 20
    TpoDocConocEmbarque = 21
    TpoDocFormSUNAT = 22
    TpoDocPolizaAdjud = 23
    TpoDocNotaCobranza = 24
    TpoDocDocATributacion = 25
    TpoDocSegComTrabRies = 27
    TpoDocProcAdquisicion = 30
    TpoDocCotizProveedor = 31
    TpoDocCuadroComparativo = 32
    TpoDocOrdenCompra = 33
    TpoDocOrdenServicio = 35
    TpoDocSolCotizacion = 36
    TpoDocNotaPedido = 41
    TpoDocNotaIngreso = 42
    TpoDocValeSalida = 43
    TpoDocRecEgreso = 44
    TpoDocHabEfectivo = 45
    TpoDocCarta = 46
    TpoDocCheque = 47
    TpoDocOrdenPago = 48
    TpoDocAutorizaEgreso = 49
    TpoDocPolizaDecUnicImport = 50
    TpoDocPolizaDUIFrac = 51
    TpoDocDeclaracionSimp = 52
    TpoDocDeclaraMensajCourier = 53
    TpoDocLiquidCobranza = 54
    TpoDocBolDep = 55
    TpoDocVoucherEgreso = 56
    TpoDocNotaCargo = 57
    TpoDocNotaAbono = 58
    TpoDocRecibosDiversos = 59
    TpoDocCompPagoDerechos = 60
    TpoDocRecArendirCuenta = 61
    TpoDocRecViaticosARendirCuenta = 62
    TpoDocDocCMAC = 63
    TpoDocPagoDietaDirectores = 64
    TpoDocCompServTranspValores = 65
    TpoDocSalidaEfectivoBoveda = 66
    TpoDocAlmacenRequerimiento = 70
    TpoDocAlmacenGuiadeRemision = 71
    TpoDocAlmacenComprdeRetencion6 = 73
    TpoDocLetras = 77
    
    TpoDocCompNoDomiciliados = 91
    TpoDocNotaCreditoEmiNoDomicilia = 97
    TpoDocNotaDebitoEmiNoDomicilia = 98
    TpoDocCartaFianza = 99
End Enum

Public Enum TipoCalculoInteres
    TpoCalcIntSimple = 0
    TpoCalcIntCompuesto = 1
End Enum

Public Enum ConstSistemas
    gConstSistAlgoritmo = 1
    gConstSistVarVerAsiento = 2
    gConstSistCodCMAC = 3
    gConstSistRutaBackup = 4
    gConstSistNombreAbrevCMAC = 5
    gConstSistNombreModulo = 6
    gConstSistCierreDiaCapta = 7
    gConstSistCierreDiaCont = 8
    gConstSistCierreAnualCont = 9
    gConstSistCierreMensualCont = 10
    gConstSistCierreDiaColoc = 11
    gConstSistCierreDiaColocPrend = 12
    gConstSistCierreSistema = 13
    gConstSistCierreMesNegocio = 14
    gConstSistFechaInicioDia = 15
    gConstSistFechaSistema = 16
    gConstSistDiasTipoCambioMant = 17
    gConstSistCtaFondoFijoME = 18
    gConstSistCtaFondoFijoMN = 19
    gConstSistCtaParaSobrantePrendario = 20
    gConstSistCtaConversionMEDol = 21
    gConstSistCtaConversiónMESoles = 22
    gConstSistTipoConverión = 23
    gConstSistNroLineasPagina = 24
    gConstSistNroLineasOrdenPago = 25
    gConstSistMargenDerCartas = 26
    gConstSistMagenIzqCartas = 27
    gConstSistMargenSupCartas = 28
    gConstSistUltVentaBarrasOro = 29
    gConstSistAgeRealizaRemate = 30
    gConstSistAgeCobJudicial = 31
    gConstSistUltCierreJud = 32
    gConstSistMinMargenRegAsist = 33
    gConstSistContraerListaOpe = 34
    gConstSistDigSubTipoFteFinan = 35
    gConstSistUltActSaldos = 36
    gConstSistDominio = 37
    gConstSistPDC = 38
    
    gConstCtaIGV = 40
    
    gConstPersCodCMACT = 41
    gConstSistUltDctoInact = 42
    gConstSistServCentralRiesgos = 43
    gConstSistValorOroDolares = 44
    gConstSistRHCargoCodApoderados = 46
    
    
    
    gConstSistRutaServidorActualiza = 47
    gConstSistRutaLocalActualiza = 48
    gConstSistVariableActualiza = 49
    
    gConstSistCMACRuc = 50
    gConstSistBitCentral = 51
    gConstSistCMACNombreCompleto = 52
    gConstSistCMACDireccion = 53
    gConstSistBitTCPonderado = 54
    gConstSistBitCtrCtasInmovil = 55
    gConstSistRHCargoCodAnalistas = 57
    
    gConstSistBitIGVxActivoCredFiscal = 60
    
    gConstSistCtaAbonoSobMN = 61
    gConstSistCtaAbonoSobME = 62
    
    gConstSistBoletaRRHH = 63
    
    gConstSistDestinoIGVDefecto = 64
    
    gConstSistCtaImpGralVentas = 65
    
    gConstSistBitRetencion6Porcent = 66
    gConstSistCtaRetencion6Porcent = 67
    gConstSistTopeRetencion6Porcent = 68
    gConstSistConversionRetencion6Porcent = 69
    gConstSistCambioRetencion6Porcent = 70
    gConstSistAgenciaEspecial = 71
    gConstSistRutaIcono = 72
    gConstSistServFSD = 77
    gConstSistVerificaRegistroEfectivo = 78
    gConstSistPigJoyasDet = 109
    gConstSistPigRemateCadaAg = 121

    ' Para Codigos antiguos
    gConstSistCodAntAhorro = 170
    gConstSistCodAntPF = 171
    gConstSistCodAntCTS = 172
    gConstSistCodAntPrendario = 173
    gConstSistCodAntOtros = 174
    gConstSistCodAntTarjeta = 175
    
    'Cuentas de Perdida y Ganacia pro tipo de cambio
    gConstSistCtaAjustTCPerdida = 111
    gConstSistCtaAjustTCGanancia = 112
    
    'Control ARendir
    gConstSistBitControARendir = 113
    
    'Control Cuentas Contables UltNivel
    gConstSistBitControCtasContUltNiv = 114
    
    'Pendientes de RRHH Para a generacion de asiento de la planilla
    gConstSistCtaPendRHSoles = 160
    gConstSistCtaPendRHDolares = 161
    gConstSistCtaPendRHCredOtrosSoles = 162
    gConstSistCtaPendRHCredOtrosDolares = 163
    
    gConstSistCtaOtrosIngresos = 164
    gConstSistCtaOtrosEgresos = 165
End Enum

Public Enum MotivoNotaAbonoCargo
    gNARecHum = 101
    gNAFoncodes = 102
    gNASedalib = 103
    gNAHidrandina = 104
    gNARegCMACRec = 105
    gNATransferencia = 106
    gNARegCargoIndebido = 107
    gNASobRemate = 108
    gNAPrestAdm = 109
    gNARegOtraAgencia = 110
    gNAConsumoLuz = 111
    gNAPagoProveedor = 112
    gNADevFondoGar = 113
    gNARendirCuenta = 114
    gNCRecHum = 201
    gNCFoncodes = 202
    gNCSedalib = 203
    gNCHidrandina = 204
    gNCComOPDevuelta = 205
    gNCComChqOtraPlaza = 206
    gNCCanjeOPEmiChqGer = 207
    gNCTransferencia = 208
    gNCChqDevuelto = 209
    gNCRetJudicial = 210
    gNCRegCMACRec = 211
    gNCDescChqDevuelto = 212
    gNCRegOtraAgencia = 213
    gNCComEmiChqGer = 214
    gNCComChqDevuelto = 215
    gNCRegAboIndebido = 216
    gNCDupTarjMag = 217
    gNCComTransferencia = 218
    gNCRendirCuenta = 219
End Enum

Public Enum MuestraIF
    MuestraInstituciones = 1
    MuestraCuentas = 7
End Enum

Public Enum TipoDiaHabil
    gTpoDiaHabPrimerDiaMes = 1
    gTpoDiaHabPenultimoDiaMes = 2
    gTpoDiaHabUltimoDiaMes = 3
End Enum

Public Enum TipoCierreRealizado
    gTpoCierreDiario = 1
    gTpoCierreMensual = 2
End Enum


Public Enum TipoPeriferico
    gPerifPINPAD = 1
    gPerifPENWARE = 2
End Enum


Public Enum TipoPinPad
    gPerifPINPAdVrfSC552 = 11
    gPerifPINPAdVrf5000 = 12
    gPerifPINPAdUnisys = 21
    gPerifPINPAdHiperCom = 31
End Enum


Public Enum ResultVerificacionTarjeta
    gClaveValida = 1
    gClaveNOValida = 2
    gTarjNoRegistrada = 3
End Enum


Public Enum TipoPuertoSerial
    gPuertoSerialCOM1 = 1
    gPuertoSerialCOM2 = 2
End Enum

Public Enum PersEstLavDinero
    gPersLavDinSinControl = 1
    gPersLavDinConControl = 2
End Enum

Public Enum AreasEstado
    AreasEstadoInActivo = 0
    AreasEstadoActivo = 1
    AreasEstadoEliminado = 2
End Enum

Public Enum AgenciasEstado
    AgenciaEstadoInActivo = 0
    AgenciaEstadoActivo = 1
    AgenciaEstadoEliminado = 2
End Enum

Public Function GeneraTablaConstante()

End Function


