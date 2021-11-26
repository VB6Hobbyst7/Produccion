VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMValores"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CaptacEstado
    gCapEstActiva = 1000        'Activa
    gCapEstBloqTotal = 1100     'Bloqueo Total
    gCapEstBloqRetiro = 1200    'Bloqueo Retiro
    gCapEstAnulada = 1300       'Anulada
    gCapEstCancelada = 1400     'Cancelada
End Enum

Public Enum CaptacTipoEstado
    gCapTpoEstActiva = 0
    gCapTpoEstInactiva = 1
    gCapTpoEstInmovilizada = 2
End Enum

Public Enum CaptacTipoTasa
    gCapTasaNormal = 100    'Captación Tasa de Interes Normal
    gCapTasaPreferencial = 101 'Captación Tasa de Interes PreferenCial
    gCapTasaEspecial = 102
End Enum

Public Enum CaptacPFFormaRetiro
    gCapPFFormRetMensual = 1
    gCapPFFormRetFinalPlazo = 2
    gCapPFFormRetLibre = 3
    gCapPFFormRetAdelantado = 4
End Enum

Public Enum CaptacOrdPagoEstado
    gCapOPEstEmitida = 1
    gCapOPEstCertifiCada = 2
    gCapOPEstCobrada = 3
    gCapOPEstRechazada = 4
    gCapOPEstExtraviada = 5
    gCapOPEstExtornada = 6
    gCapOPEstAnulada = 7
End Enum

Public Enum CaptacRelacPersona
    gCapRelPersTitular = 10
    gCapRelPersApoderado = 11
    gCapRelPersRepTitular = 12
    gCapRelPersRepSuplente = 13
    gCapRelPersPromotor = 14
End Enum

Public Enum CaptacTarjetaEstado
    gCapTarjEstRegistrada = 0   'Estado Tarjeta Registrada
    gCapTarjEstActiva = 1   'A,N,K   Estado Tarjeta MagnetiCa Activa
    gCapTarjEstBloqueada = 2    'B   Estado Tarjeta Magnetica Bloqueada
    gCapTarjEstCancelada = 3    'C,V   Estado Tarjeta Magnetica Cancelada
    gCapTarjEstCmbioClave = 4
End Enum

Public Enum CaptacMotBloqueoRet
    gCapMotBlqRetMandatoJudicial = 1
    gCapMotBlqRetOrdenTitular = 2
    gCapMotBlqRetGarantia = 3
    gCapMotBlqRetFaltanFirmas = 4
    gCapMotBlqRetFaltanDocumentos = 5
    gCapMotBlqRetCuentaInactiva = 6
    gCapMotBlqRetRobo = 7
    gCapMotBlqRetOrdenAdministracion = 8
    gCapMotBlqRetDepositoIndebido = 9
    gCapMotBlqRetOrdenPagoRechazada = 10
    gCapMotBlqRetOrdenPagoDevuelta = 11
    gCapMotBlqRetCambioEstructuraPasivo = 24 'FRHU 20140301 RQ14008
End Enum
    
Public Enum CaptacMotBloqueoTot
    gCapMotBlqTotMandatoJudicial = 1
    gCapMotBlqTotOrdenTitular = 2
    gCapMotBlqTotGarantia = 3
    gCapMotBlqTotFaltanFirmas = 4
    gCapMotBlqTotFaltanDocumentos = 5
    gCapMotBlqTotCuentaInactiva = 6
    gCapMotBlqTotRobo = 7
    gCapMotBlqTotOrdenAdministracion = 8
    gCapMotBlqTotDepositoIndebido = 9
    gCapMotBlqTotOrdenPagoRechazada = 10
    gCapMotBlqTotOrdenPagoDevuelta = 11
    gCapMotBlqTotSobRemate = 12
End Enum

Public Enum CaptacMotBloqueoParcial
    gCapMotBlqParMandatoJudicial = 1
    gCapMotBlqParOrdenTitular = 2
    gCapMotBlqParGarantia = 3
    gCapMotBlqParOrdenAdministracion = 4
    gCapMotBlqParDepositoIndebido = 5
    gCapMotBlqParOrdenPagoRechazada = 6
    gCapMotBlqTotOrdenPagoDevuelta = 7
End Enum

Public Enum CaptacTipoBloqueo
    gCapTpoBlqTotal = 11
    gCapTpoBlqRetiro = 12
    gCapTpoBlqParcial = 13
End Enum





Public Enum CaptacParametro
    gCostoChqME = 2001
    gCostoChqMN = 2002
    gCostoOPRechME = 2003
    gCostoOPRechMN = 2004
    gDctoExtME = 2005
    gDctoExtMN = 2006
    
    gDctoExtMExPag = 2051  ' ppoa
    gDctoExtMNxPag = 2052  ' ppoa
    
    gDiasMaxOPCert = 2007
    gDiasMinValChq = 2008
    gMinApeAhoME = 2009
    gMinApeAhoMN = 2010
    gMinApeAhOPPJME = 2011
    gMinApeAhOPPJMN = 2012
    gMinApeAhOPPNME = 2013
    gMinApeAhOPPNMN = 2014
    gMinApePFME = 2015
    gMinApePFMN = 2016
    gMonDescInacME = 2017
    gMonDescInacMN = 2018
    gMonFSD = 2019
    gNumOrdPagTal = 2020
    gPorRetCTS = 2021
    gSaldMinAhoME = 2022
    gSaldMinAhoMN = 2023
    gCtaBacoOrdPag = 2024
    gDiasMaxInactCond1 = 2025
    gDiasMaxInactCond2 = 2026
    gSaldMinMNInactCond2 = 2027
    gSaldMinMEInactCond2 = 2028
    gNumVecesMinRechOP = 2029
    gMonDctoMNRechOP = 2030
    gMonDctoMERechOP = 2031
    gMonOpeLavDineroME = 2032
    gMonMensLavDineroME = 2033
    gCostoDuplicadoTarjeta = 2034
    gCostoVentaBases = 2035
    gCostoConstanciaCancCred = 2036
    gCostoElaboracionContrato = 2037
    gVisualizaExtornoExtracto = 2044
    gCostoComServPublico = 2045
    gCostoComServSATInfraccion = 2046
    gCostoComServSATRecibosDerechos = 2047
    gMontoSobranteCajero = 2048
    gMontoFaltanteCajero = 2049
    gCostoOperacionCMACLlam = 2050
    gMontoMNx31Ope = 2082
    gMontoMEx31Ope = 2083
    
    gMontoEurox31Ope = 2102 '*** PEAC 20090727
    
    gComisionRetOtraAge = 2084
        
    gMontoLimiMNRetOtraAge = 2085
    gMontoLimiMERetOtraAge = 2086
    gNroMaxRet = 2087
    gPlazoMinCancelacion = 2088
    gPlazoMinCancelNañito = 2089
    gPlazoMinCancelPandero = 2090
    gPorRetiroAdicionalCTS = 2099 'By Capi Agosto 2007 (Acta 014-2007)
    'By Capi 28022008
    gMonOpeWestLavDineroME = 2100
    gMonMulWestLavDineroME = 2101
    gCtaAfeWestOpeLavDineroME = 2102
    gCtaAfeWestOpeLavDineroMN = 2103
    gMonOpeNoWestLavDineroME = 2104
    gMonMulNoWestLavDineroME = 2105
    gMonOpeNoWestLavDineroME02 = 2109 'madm 20100722
    gMonOpeNoWestLavDineroME0725 = 2111 'madm 20100722
    
End Enum

Public Enum CaptacOperacion
    'AHORROS
    gAhoApeEfec = 200101    'Apertura Efectivo
    gAhoApeChq = 200102     'Apertura Cheque
    gAhoApeTransf = 200103      'Apertura Transferencia
    gAhoApeLoteEfec = 200104
    gAhoApeLoteChq = 200105
    gAhoApeSobRemate = 200106
    gAhoApeCargoCta = 200109 'Apertura con Cargo a Cuenta 'JUEZ 20131211
    gAhoApeLoteTransfBanco = 200110 'Apertura Lote Transferencia Banco ERS017
    gAhoDepEfec = 200201    'Depósito Efectivo
    gAhoDepChq = 200202     'Depósito Cheque
    gAhoDepTransf = 200203      'Depósito Transferencia
    gAhoDepEntConv = 200204 'Depósito Entidad Convenio
    gAhoDepGratRRHH = 200205  'Depósito Gratificación RRHH
    gAhoDepPlanRRHH = 200206  'Depósito Planilla RRHH
    gAhoDepPagProvBien = 200207  'Depósito Pago Prov. de Bienes
    gAhoDepPagProvServ = 200208  'Depósito Pago. Prov. de Servicios
    gAhoDepOtrosIngRRHH = 200209  'Depósito Otros Ingresos RRHH
    gAhoDepPagFoncodes = 200210  'Depósito Pagos FONCODES
    gAhoDepPagServHidrandina = 200211  'Depósito Pago. Servicios Hidrandina
    gAhoDepPagServSedalib = 200212  'Depósito Pago. Servicios Sedalib
    gAhoDepSobRemate = 200213  'Depósito Sobrante de Remate
    gAhoDepDevFondoGar = 200214  'Depósito Devolución Fondo Garantía
    gAhoDepSobCaja = 200215  'Depósito Sobrantes de Caja
    gAhoDepPagFidCapital = 200216  'Deposito Pagos Fideicomiso - Capital
    gAhoDepPagFidInteres = 200217  'Deposito Pagos Fideicomiso - Interes
    gAhoDepPagFidGastos = 200218   'Deposito Pagos Fideicomiso - Gastos
    gAhoDepPagServEdelnor = 200219  'Depósito Pago. Servicios Edelnor
    gAhoDepAboIntPF = 200220  'Depósito Abono Int. Plazo Fijo
    gAhoDepAboOtrosConceptos = 200221  'Depósito Abono Otros Conceptos
    gAhoDepPlanRRHHAdelantoSueldos = 200222 'Deposito de Adelanto de Sueldos
    gAhoDepRegCMACLlam = 200240  'Depósito Abono Regularización Llamada
    gAhoDepDevCredPersonales = 200241  'Depósito Devolución Creditos Personales
    gAhoDepLeasing = 200260 'Depósito Leasing
    gAhoDepCtaRecaudoEcotaxi = 200261 'EJVG20120518
    gAhoDepCtaxPenalidadEcotaxixDiasRetrasoEntrega = 200262 'EJVG20120622
    gAhoDepDirectoClub = 200264 'RIRO20131102 ERS145
    'RIRO ERS017 ***
    gAhoDepositoHaberesEnLoteEfec = 200266
    gAhoDepositoHaberesEnLoteTransf = 200267
    'END RIRO ******
    gAhoDepAfilSegTarjeta = 200268 'JUEZ 20140520
    gAhoDepActivacionSegTarjeta = 200269 'FRHU 20140602 ERS068-2014
    gAhoDepositoHaberesEnLoteChq = 200282 'CTI620210506
    
    gAhoRetEfec = 200301    'Retiro Efectivo
    gAhoRetOP = 200302      'Retiro Orden Pago
    gAhoRetTransf = 200303      'Retiro Nota Cargo
    gAhoRetOPCanje = 200304 'Retiro Orden Pago Canje
    gAhoRetOPCert = 200305  'Retiro Orden Pago Certificada
    gAhoRetOPCertCanje = 200306  'Retiro Orden Pago Certificada Canje
    gAhoRetFondoFijo = 200307 'Retiro Fondo Fijo
    gAhoRetFondoFijoCanje = 200308 'Retiro Fondo Fijo Canje
    gAhoRetAnulChq = 200309 'Retiro Anulación Cheque
    gAhoRetEmiChq = 200310 'Retiro por emisión de Cheque (Simple, Gerencia)
    gAhoRetEmiChqCanjeOP = 200311 'Retiro por emisión de Cheque (Simple, Gerencia) con canje OP
    gAhoRetOtrosDctosRRHH = 200312 'Retiro Otros Descuentos RRHH
    gAhoRetComFoncodes = 200313 'Retiro Comisión Cobranza FONCODES
    gAhoRetComServHidrandina = 200314 'Retiro Comisión Cobranza Hidrandina
    gAhoRetComServSEDALIB = 200315 'Retiro Comisión Cobranza SEDALIB
    gAhoRetRetencionJudicial = 200316 'Retiro Retención Judicial
    gAhoRetDuplicadoTarj = 200317 'Retiro Duplicado de Tarjeta
    gAhoRetComOrdPagDev = 200318 'Retiro Comisión Orden Pago Devuelta
    gAhoRetComChqOtraPlaza = 200319 'Retiro Comisión Cheque Otra Plaza
    gAhoRetChequeDevuelto = 200320 'Retiro Cheque Devuelto
    gAhoRetComEmiChq = 200321 'Retiro Comisión Emisión Cheque Simple/Gerencia
    gAhoRetComChqDevuelto = 200322 'Retiro Comisión Cheque Devuelto
    gAhoRetComTransferencia = 200323 'Retiro Comisión Transferemcia
    gAhoRetComServFideicomiso = 200324  'Retiro Comisión Cobranza Fideicomiso
    gAhoRetComServEDELNOR = 200325 'Retiro Comisión Cobranza EDELNOR
    gAhoRetConsultaSaldos = 200326 'Retiro Consulta de Saldos
    gAhoRetPorteCargoCuenta = 200327 'Retiro Porte Cargo a cuenta
    gAhoRetComVentaBases = 200328 'Retiro Comisión Venta de Bases
    gAhoRetOtrosConceptos = 200329 'Retiro Otros Conceptos
    gAhoRetRegCMACLlam = 200330 'Retiro Regularización CMAC Llamada
    gAhoRetComision = 200358 'Cobro de Comision por Retiro en OtraAgencias y por max de Ope.
    gAhoDepComision = 200359 'BRGO 20110127 - Cobro de Comision por Depósitos en OtraAgencias y por max de Ope.
    gAhoCargoCtaxPenalidadEcotaxixVehiculoDistintoAcordado = 200363 'EJVG20120622
    gAhoCargoCtaxPenalidadEcotaxixDiasRetrasoEntrega = 200364 'EJVG20120622
    gAhoCargoCobroComDiversasAho = 200366 'JUEZ 20130831
    gAhoCargoCtaApert = 200367 'JUEZ 20131212
    gAhoCargoCtaAumCapPF = 200368 'JUEZ 20131218
    gAhoCargoCtaPagoCred = 200369 'JUEZ 20131227
    gAhoCargoAfilSegTarjeta = 200374 'JUEZ 20140520
    gAhoCargoCtaPagoServ = 200377 'JUEZ 20150130
    gAhoCargoComisionServDebitoAuto = 200378 'JUEZ 20150130
    gAhoCargoAfilSegSepelio = 200380 'RECO20151123 ERS073-2015
    gAhoRetComOPCanje = 200399 'MADM 20101115 - Comision Retiro Orden Pago Canje
    
    gAhoCargoCtaPagoCredAuto = 200376 'WIOR 20140620
    
    gAhoCargoCtaCancelaPigno = 203010 'CTI4 ERS0112020
    gAhoCargoCtaCancelaPignoNor = 203011 'CTI4 ERS0112020
    gAhoCargoCtaCancelaPignoMor = 203012 'CTI4 ERS0112020
    gAhoCargoCtaAmortizaPignoNor = 203020 'CTI4 ERS0112020
    gAhoCargoCtaRecupCoAdj = 203030 'CTI4 ERS0112020
    gAhoCargoCtaRenovaPigno = 203040 'CTI4 ERS0112020
    gAhoCargoCtaRenovaPignoNor = 203041 'CTI4 ERS0112020
    gAhoCargoCtaRenovaPignoMor = 203042 'CTI4 ERS0112020
    gAhoCargoCtaDepositoLoteCTS = 203050 'CTI4 ERS0112020
    
    gAhoCancAct = 200401 'Cancelación Activa
    gAhoCancInact = 200402 'Cancelación Inactiva
    gAhoCancTransfAct = 200403 'Cancelación Activa
    gAhoCancTransfInact = 200404 'Cancelación Inactiva
    gAhoCancInmovilizada = 200405 'Cancelacion Inmovilizada
    gAhoCancInmovFSD = 200406 'Cancelación Imnmovilizada pase FSD
    gAhoCancSobRemate = 200407  'Cancelacion por Sobrante de Remate
    
    gAhoCargoCompra = 203910 'CTI6 ERS0112020
    gAhoCargoVenta = 203911 'CTI6 ERS0112020
    gAhoCargoServicioRecaudo = 203912 'CTI6 ERS0112020
    
    gAhoCargoDineroElectronico = 203913 'CTI6 ERS0112020
    gAhoDepositoDineroElectronico = 203914 'CTI6 ERS0112020
    
    gAhoCancSobregiroOP = 200408  'Cancelacion por Sobregiro de ordenes de Pago
    gAhoCancTransfAbCtaBco = 200411 'RIRO20131226 ERS137
        
    gAhoCapIntAct = 200501  'Capitalización Intereses Activa
    gAhoCapIntInact = 200502  'Capitalización Intereses Inactiva
    gAhoCapIntMig = 200504 'Capitalización Intereses de Migración--gregado por ELRO el 20130326, según TI-ERS011-2013
    gAhoDctoEmiExt = 200601 'Descuento Emisión Extracto
    gAhoDctoEmiOP = 200602  'Descuento Emisión Chequera
    gAhoDctoInactiva = 200603  'Descuento Cuenta Inactiva
    gAhoDctoEmiExtDebitoAuto = 200605 'Descuento Emisión Extracto Debito Automatico 'JUEZ 20150130
    gAhoOPEmision = 200701 'Orden Pago Emisión
    gAhoOPAnulacion = 200702  'Orden Pago Anulación
    gAhoOPCertificacion = 200703  'Orden Pago Certificación
    gAhoOPDevuelta = 200704  'Orden Pago Devuelta
    gAhoOPDevueltaCanje = 200705  'Orden Pago Devuelta Canje
    gAhoEstActInac = 200801      'Activa -Inactiva
    gAhoEstInacAct = 200802      'Inactiva -Activa
    gAhoEstInacInmov = 200803   'Inactiva -Inmovilizada
    gAhoEstInmovAct = 200804      'Inmovilizada -Activa
    gAhoTransferencia = 200900      'Transferencias
    gAhoTransAbono = 200901  'Transferencia Abono
    gAhoTransCargo = 200902  'Transferencia Cargo
    
    'Add By GITU 23-10-2012
    gAhoTransAbonoL = 200904  'Transferencia Abono en Lote
    gAhoTransCargoL = 200905  'Transferencia Cargo en Lote
    'End GITU
    
    gAhoTransAboCargIndebido = 200906  'Transferencia por Abono/Cargo Indebido
    gAhoTransAboIndebMNMN = 200907  'Transferencia MN - MN
    gAhoTransAboIndebMEME = 200908  'Transferencia ME - ME
    gAhoTransCargoIndeb = 200910  'Transferencia Cargo Indebido
    
    gAhoTransCargoLC = 200911  'Transferencia Cargo en Lote Convenio, RIRO 20150612 ERS146 ADD
    
    gAhoConsSaldo = 201001  'Consulta Saldo
    gAhoConsMovimiento = 201002 'Consulta Movimientos
    
    gAhoMigracion = 202001 'Migracion de Ahorros--gregado por ELRO el 20130326, según TI-ERS011-2013
     
    'Plazo Fijo
    gPFApeEfec = 210101  'Apertura Efectivo
    gPFApeChq = 210102  'Apertura Cheque
    gPFApeTransf = 210103  'Apertura Transferencia
    gPFApeLoteEfec = 210104
    gPFApeLoteChq = 210105
    gPFApeCargoCta = 210107 'Apertura con Cargo a Cuenta 'JUEZ 20131212
    gPFApeLoteTransf = 210109 'CTI7 OPEv2***************************************
    gPFRetInt = 210201  'Retiro de Intereses Efectivo
    gPFRetIntAboAho = 210202  'Retiro de Intereses Abono Cta Ahorros
    gPFRetRegCapital = 210204 'Retiro Regularización de Capital
    gPFRetIntAdelantado = 210206  'Retiro de Intereses Adelantado
    gPFCancEfec = 210301  'Cancelación
    gPFCancTransf = 210302  'Cancelación Transferencia
    gPFCapInt = 210401  'Capitalización Intereses
    gPFRenovacion = 210501  'Renovación
    gPFDupCert = 210601 'Duplicado de Certificado de Plazo Fijo
    gPFBusqCredOend = 210602 'Búsqueda de Crédito Pendiente
    gPFConsSaldo = 210701 'Consulta de Saldos
    gPFConsMovimiento = 210702 'Ultimos Movimientos
    gPFAumCapEfec = 210801 'Aumento de Capital Efectivo
    gPFAumCapchq = 210802 'Aumento de Capital Cheque
    gPFAumCapTasaPactEfec = 210803 'Aumento de Capital Efectivo Tasa Pactada
    gPFAumCapTasaPactChq = 210804  'Aumento de Capital Cheque Tasa Pactada
    gPFDismCapEfec = 210805 'Disminución de Capital Efectivo
    gPFAumDismCapInt = 210806 'Aum/Dism Capitalizacion Intereses
    gPFAumCapTrans = 210807 'Aumento de Capital Efectivo Transferencia a Banco
    gPFAumCapTasaPactTrans = 210808 'Aumento de Capital Efectivo Transferencia a Banco Tasa Pactada
    gPFAumCapCargoCta = 210813 'Aumento de Capital con Cargo a Cuenta 'JUEZ 20131212
    gPFRetIntAboCtaBanco = 210209  'Retiro de Intereses Abono Cta Banco RIRO20131212 ERS137
    
    'CTS
    gCTSApeEfec = 220101  'Apertura Efectivo
    gCTSApeChq = 220102  'Apertura Cheque
    gCTSApeTransf = 220103  'Apertura Transferencia
    gCTSApeLoteEfec = 220104
    gCTSApeLoteChq = 220105
    gCTSApeLoteTransf = 220106 'Apertura Lote Nota Abono'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSApeLoteTransfNew = 220108 'Apertura Lote Nota Abono CTS'***Agregado por CTI7 el 20210618, según OpeV2
    gCTSDepEfec = 220201    'Depósito Efectivo
    gCTSDepChq = 220202     'Depósito Cheque
    gCTSDepTransf = 220203      'Depósito Nota Abono
    gCTSDepPlanRRHH = 220204      'Depósito Planilla RRHH
    gCTSDepAboOtrosConceptos = 220205 'By capi 05032009 Acta 025-2009
    gCTSDepAfilSegTarjeta = 220212 'JUEZ 20140520
    gCTSRetEfec = 220301      'Retiro Efectivo
    gCTSRetTransf = 220302  'Retiro Nota Cargo
    gCTSCargoCobroComDiversasAho = 220305 'JUEZ 20130831
    gCTSCargoAfilSegTarjeta = 220310 'JUEZ 20140520
    gCTSCargoAfilSegSepelio = 220315 'RECO20151123 ERS073-2015
    gCTSDctoEmiExt = 220317   'JUEZ 20151229
    gCTSCancEfec = 220401      'Cancelación Efectivo
    gCTSCancTransf = 220402      'Cancelación Transferencia
    gCTSCancTransfBco = 220403      'Cancelación Transferencia RIRO20131212 ERS137
    gCTSCapInt = 220501  'Capitalización Intereses
    gCTSConsSaldo = 220601 'Consulta de Saldos
    gCTSConsMovimiento = 220602 'Ultimos Movimientos
    gCTSDepLotEfec = 220207    'Depósito Lote Efectivo'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSDepLotChq = 220208     'Depósito Lote Cheque'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSDepLotTransf = 220209      'Depósito Lote Nota Abono'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSPasarNoActiva = 220803 'Pasado a CTS No Activo'***Agregado por ELRO el 20121123, según OYP-RFC101-2012
    gCTSDepLotCargoCta = 220215 'CTI4ERS0112020 Deposito Lote Cargo a Cuenta
    
    gAhoExt = 230000  'EXTORNOS AHORROS
    gAhoExtApeEfec = 230101  'Extorno Apertura Efectivo
    gAhoExtApeChq = 230102  'Extorno Apertura Cheque
    gAhoExtApeTransf = 230103  'Extorno Apertura Transferencia
    gAhoExtApeLoteEfec = 230104  'Extorno Apertura Lote Cheque
    gAhoExtApeLoteChq = 230105  'Extorno Apertura Lote Efectivo
    gAhoExtApeCargoCta = 230106 'Extorno Apertura Cargo a Cuenta 'JUEZ 20131226
    gAhoExtApeLoteTransfBanco = 230107 'Extorno Apertura Lote Transferencia 'RIRO 20140528
    gAhoExtDepEfec = 230201 'Extorno Depósito Efectivo
    gAhoExtDepChq = 230202  'Extorno Depósito Cheque
    gAhoExtDepTransf = 230203  'Extorno Depósito Transferencia
    
    gAhoExtDepGratRRHH = 230205   'Extorno Depósito Gratificación RRHH
    gAhoExtDepPlanRRHH = 230206   'Extorno Depósito Planilla RRHH
    gAhoExtDepPagProvBien = 230207  'Extorno Depósito Pago Prov. de Bienes
    gAhoExtDepPagProvServ = 230208  'Extorno Depósito Pago. Prov. de Servicios
    gAhoExtDepOtrosIngRRHH = 230209  'Extorno Depósito Otros Ingresos RRHH
    gAhoExtDepPagFoncodes = 230210  'Extorno Depósito Pagos FONCODES
    gAhoExtDepPagServHidrandina = 230211  'Extorno Depósito Pago. Servicios Hidrandina
    gAhoExtDepPagServSedalib = 230212  'Extorno Depósito Pago. Servicios Sedalib
    gAhoExtDepSobRemate = 230213  'Extorno Depósito Sobrante de Remate
    gAhoExtDepDevFondoGar = 230214  'Extorno Depósito Devolución Fondo Garantía
    gAhoExtDepSobCaja = 230215  'Extorno Depósito Sobrantes de Caja
    gAhoExtDepPagFidCapital = 230216  'Extorno Deposito Pagos Fideicomiso - Capital
    gAhoExtDepPagFidInteres = 230217  'Extorno Deposito Pagos Fideicomiso - Interes
    gAhoExtDepPagFidGastos = 230218   'Extorno Deposito Pagos Fideicomiso - Gastos
    gAhoExtDepPagServEdelnor = 230219  'Extorno Depósito Pago. Servicios Edelnor
    gAhoExtDepAboIntPF = 230220  'Extorno Depósito Intereses Plazo Fijo
    gAhoExtDepOtrosConceptos = 230222  'Extorno Depósito Intereses Plazo Fijo
    gAhoExtDepRegCMACLlam = 230223  'Extorno Depósito Regularización Llamada
    gAhoExtDepDevCredPersonales = 230223  'Extorno Depósito Devolución Créditos Personales
    gAhoExtDepComision = 230225 'Extorno de Comisiones
    gAhoExtDepPagFocmacm = 230226 'Extorno Deposito Pagos FOCMACM 'FRHU 20150520 ERS022-2015
    gAhoExtDirectoClub = 230264 'Extorno Deposito Club Trabajadores. RIRO20131102 ERS145
    gAhoExtDepositoHaberesEnLoteEfec = 230265 ' RIRO20140528 ERS017
    gAhoExtDepositoHaberesEnLoteTransf = 230266 ' RIRO20140528 ERS017
    gAhoExtDepositoHaberesEnLoteChq = 230267 ' RIRO20140528 ERS017
    
    gAhoExtRetEfec = 230301  'Extorno Retiro Efectivo
    gAhoExtRetOP = 230302  'Extorno Retiro Orden Pago
    gAhoExtRetTransf = 230303  'Extorno Retiro Transferencia
    gAhoExtRetOPCanje = 230304  'Extorno Retiro Orden Pago Canje
    gAhoExtRetOPCert = 230305  'Extorno Retiro Orden Pago Certificada
    gAhoExtRetOPCertCanje = 230306  'Extorno Retiro Orden Pago Certificada Canje
    gAhoExtRetFondoFijo = 230307 'Extorno Retiro Fondo Fijo
    gAhoExtRetFondoFijoCanje = 230308 'Extorno Retiro Fondo Fijo Canje
    gAhoExtRetAnulChq = 230309  'Extorno Retiro Anulación Cheque
    gAhoExtRetEmiChq = 230310 'Extorno Retiro Emisión Cheque (Simple, Gerencia)
    gAhoExtRetEmiChqCanjeOP = 230311 'Extorno Emisión Cheque Canje OP
    gAhoExtRetOtrosDctosRRHH = 230312 'Retiro Otros Descuentos RRHH
    gAhoExtRetComFoncodes = 230313 'Retiro Comisión Cobranza FONCODES
    gAhoExtRetComServHidrandina = 230314 'Retiro Comisión Cobranza Hidrandina
    gAhoExtRetComServSEDALIB = 230315 'Retiro Comisión Cobranza SEDALIB
    gAhoExtRetRetencionJudicial = 230316 'Retiro Retención Judicial
    gAhoExtRetDuplicadoTarj = 230317 'Retiro Duplicado de Tarjeta
    gAhoExtRetComOrdPagDev = 230318 'Retiro Comisión Orden Pago Devuelta
    gAhoExtRetComChqOtraPlaza = 230319 'Retiro Comisión Cheque Otra Plaza
    gAhoExtRetChequeDevuelto = 230320 'Retiro Cheque Devuelto
    gAhoExtRetComEmiChq = 230321 'Retiro Comisión Emisión Cheque Simple/Gerencia
    gAhoExtRetComChqDevuelto = 230322 'Retiro Comisión Cheque Devuelto
    gAhoExtRetComTransferencia = 230323 'Retiro Comisión Transferemcia
    gAhoExtRetComServEdelnor = 230324 'Retiro Comisión Cobranza Edelnor
    gAhoExtRetConsultaSaldos = 230326 'Retiro Consulta de Saldos
    gAhoExtRetPorteCargoCuenta = 230327 'Retiro Porte Cargo a cuenta
    gAhoExtRetComVentaBases = 230328 'Retiro Comisión Venta de Bases
    gAhoExtRetOtrosConceptos = 230330 'Retiro Comisión Venta de Bases
    gAhoExtRetRegCMACLlam = 230331 'Retiro Regularización CMAC Llamada
    gAhoExtretComision = 230359 'Extorno de Comisiones
    gAhoExtCancAct = 230401  'Extorno Cancelación Activa
    gAhoExtCancTransfAct = 230402  'Extorno Cancelación Transferencia
    gAhoExtCancSobRemate = 230407  'Extorno Cancelacion Sobrante Remate
    gAhoExtCanctransf = 230411  'Extorno Cancelacion Transf. Ab Cta RIRO20131212 ERS137
    gAhoExtDctoEmiExt = 230601 'Extorno Descuento Emisión Extracto
    gAhoExtDctoEmiOP = 230602 'Extorno Descuento Emisión Chequera
    gAhoExtCargoComDivAho = 230366 'Extorno cargo por cobro de comisiones diversas de ahorros 'JUEZ 20130906
    gAhoExtCargoCtaApert = 230367 'Extorno Cargo a Cuenta por Apertura 'JUEZ 20131226
    gAhoExtCargoCtaAumCapPF = 230368 'Extorno Cargo a Cuenta por Aumento Capital 'JUEZ 20131226
    gAhoExtCargoCtaPagoCred = 230369 'Extorno Cargo a Cuenta por Pago Credito 'JUEZ 20131226
    gAhoExtCargoCtaCancelPigno = 230370 'CTI4 ERS0112020
    gAhoExtCargoCtaPagoParcialPigno = 230371    'CTI4 ERS0112020
    gAhoExtCargoCtaRecuperaContratoAdjPigno = 230372    'CTI4 ERS0112020
    gAhoExtCargoCtaRenovPigno = 230373    'CTI4 ERS0112020
    gAhoExtCargoCtaCTSDepLot = 230374 'CTI4 ERS0112020
    gAhoExtCargoCtaContanciaNoAdeudo = 230378 'CTI4 ERS0112020
    gAhoExtCargoCtaComisionesDiversasCredito = 230379 'CTI4 ERS0112020
    gAhoExtTransferencia = 230900 'Extorno Transferencia entre cuentas
    gAhoExtTransAbono = 230901
    gAhoExtTransCargo = 230902
    'Add By GITU 06-11-2012
    gAhoExtTransAbonoL = 230904
    gAhoExtTransCargoL = 230905
    gAhoExtComiRetSinTarj = 230361 'Extorno de Comisiones por retiro sin tarjeta
    gAhoExtComisionCargoCtaEmisionCheque = 230362 'CTI4 ERS0112020 Extorno Comision Cargo a Cuenta por Emision de Cheque.
    gAhoExtComisionCargoCtaRetTrAboCtaBco = 230363 'CTI4 ERS0112020 Extorno Comision Cargo a Cuenta por Transferencia Abono Cta Bco.
    'End GITU
    gAhoExtMigracion = 231001 'Migracion de Ahorros--gregado por ELRO el 20130401, según TI-ERS011-2013
    
    gPFExt = 240000  'EXTORNOS PLAZO FIJO
    gPFExtApeEfec = 240101  'Extorno Apertura Efectivo
    gPFExtApeChq = 240102  'Extorno Apertura Cheque
    gPFExtApeTransf = 240103  'Extorno Apertura Transferencia
    gPFExtApeLoteEfec = 240104  'Extorno Apertura Lote Efectivo
    gPFExtApeLoteChq = 240105  'Extorno Apertura Lote Cheque
    gPFExtApeCargoCta = 240106 'Extorno Apertura Cargo a Cuenta 'JUEZ 20131226
    gPFExtApeLoteTransf = 240107  'Extorno Apertura Lote Transferencia
    gPFExtRetInt = 240201  'Extorno Retiro Intereses Efectivo
    gPFExtRetIntAboAho = 240202  'Extorno Retiro Intereses Abono Cta Ahorro
    gPFExtRetIntCash = 240203  'Extorno Retiro Intereses Cash - PEAC 20091229
    gPFExtRetIntAboCtaTransf = 240204  'Extorno Retiro Intereses Abono Cta Bancos RIRO20131212 ERS137
    gPFExtCancEfec = 240301  'Extorno Cancelación Efectivo
    gPFExtCancTransf = 240302  'Extorno Cancelación Transferencia     gPFExtCancTransf = 240302
    gPFExtCancTransfAbBco = 240304  'Extorno Cancelación Transferencia Abono Cta Bco. RIRO20131212 ERS137
    gPFExtAumCapEfec = 240801 'Extorno Aumento de Capital Efectivo
    gPFExtAumCapchq = 240802 'Extorno Aumento de Capital Cheque
    gPFExtAumCapTasaPactEfec = 240803 'Extorno Aumento de Capital Efectivo Tasa Pactada
    gPFExtAumCapTasaPactChq = 240804  'Extorno Aumento de Capital Cheque Tasa Pactada
    gPFExtDismCapEfec = 240805 'Extorno  Disminución de Capital Efectivo
    gPFExtAumCapTrans = 240807 'Extorno Aumento de Capital Efectivo Transferencia a Banco
    gPFExtAumCapTasaPactTrans = 240808 'Extorno Aumento de Capital Efectivo Transferencia a Banco Tasa Pactada
    gPFExtAumCapCargoCta = 240809 'Extorno Aumento de Capital Cargo a Cuenta
  
    
    gCTSExt = 250000  'EXTORNOS CTS
    gCTSExtApeEfec = 250101  'Extorno Apertura Efectivo
    gCTSExtApeChq = 250102  'Extorno Apertura Cheque
    gCTSExtApeTransf = 250103  'Extorno Apertura Transferencia
    gCTSExtApeLoteEfec = 250104  'Extorno Apertura Lote Efectivo
    gCTSExtApeLoteChq = 250105  'Extorno Apertura Lote Cheque
    gCTSExtApeLoteTransf = 250106 'Extorno Apertura Lote Nota Abono'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSExtDepEfec = 250201    'Extorno Depósito Efectivo
    gCTSExtDepChq = 250202     'Extorno Depósito Cheque
    gCTSExtDepTransf = 250203      'Extorno Depósito Transferencia
    gCTSExtRetEfec = 250301      'Extorno Retiro Efectivo
    gCTSExtRetTransf = 250302  'Extorno Retiro Nota Cargo
    gCTSExtRetTransfAbCta = 250306  'Extorno Retiro Abono en Cuenta RIRO20131212 ERS137
    gCTSExtCargoComDivAho = 250305 'Extorno cargo por cobro de comisiones diversas de ahorros 'JUEZ 20130906
    gCTSExtCancEfec = 250401      'Extorno Cancelación Efectivo
    gCTSExtCancTransf = 250402      'Extorno Cancelación Transferencia
    gCTSExtCancTransfAbCta = 250403      'Extorno Cancelación Transferencia Abono Cuenta RIRO20131212 ERS137
    gCTSExtDepLotEfec = 250207    'Extorno Depósito Lote Efectivo'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSExtDepLotChq = 250208     'Extorno Depósito Lote Cheque'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSExtDepLotTransf = 250209  'Extorno Depósito Lote Transferencia'***Agregado por ELRO el 20121023, según OYP-RFC101-2012
    gCTSExtDepActiInact = 250210  'Extorno Activo/Inactiva'***ALPA20130827
    gCTSExtDepLotCargoCta = 250211 'CTI4 ERS0112020
    
    'Operaciones con CMACs Recepcion
    gCMACOAAhoDepEfec = 260101
    gCMACOAAhoDepChq = 260102
    gCMACOAAhoRetEfec = 260103
    gCMACOAAhoRetOP = 260104
    gCMACOAAhoRetOPCert = 260105
    gCMACOAPFRetInt = 260201
    gCMACOACTSDepEfec = 260301
    gCMACOACTSRetEfec = 260302

    'Operaciones con CMACs Llamada
    gCMACOTAhoDepEfec = 260501
    gCMACOTAhoDepChq = 260502
    gCMACOTAhoRetEfec = 260503
    gCMACOTAhoRetOP = 260504
    
    'Extornos Operaciones con CMACs Recepcion
    gCMACOAAhoExtDepEfec = 270101
    gCMACOAAhoExtDepChq = 270102
    gCMACOAAhoExtRetEfec = 270103
    gCMACOAAhoExtRetOP = 270104
    gCMACOAAhoExtRetOPCert = 270105
    gCMACOAPFExtRetInt = 270201
    gCMACOACTSExtDepEfec = 270301
    gCMACOACTSExtRetEfec = 270302
    
    'Extornos Operaciones con CMACs Llamada
    gCMACOTAhoExtDepEfec = 270501
    gCMACOTAhoExtDepChq = 270502
    gCMACOTAhoExtRetEfec = 270503
    gCMACOTAhoExtRetOP = 270504

    'Códigos para reportes
    gCapReportes = 280000
    
    gCapRepDiario = 280100
    gCapRepDiaCap = 280200
    gCapRepDiaCapEstadAho = 280201
    gCapRepDiaCapEstadPF = 280202
    gCapRepDiaCapEstadCTS = 280203
    gCapRepDiaCapCtasMov = 280204
    gCapRepDiaCapSaldTpoCta = 280205
    gCapRepDiaCapEstratCta = 280206
    gCapRepDiaCapPFVenc = 280207
    gCapRepDiaCapInact = 280208
    gCapRepDiaCapConsInact = 280209
    gCapRepDiaCapApert = 280210
    gCapRepDiaCapCanc = 280211
    gCapRepDiaCartaRenPF = 280212
    
    gCapRepDiaServ = 280300
    gCapRepDiaServGirosApert = 280301
    gCapRepDiaServGirosCanc = 280302
    gCapRepDiaServConvCob = 280303
    
    gCapRepMensual = 280500
    gCapRepMensCap = 280600
    gCapRepMensCapSaldCtaAhoPN = 280601
    gCapRepMensCapSaldCtaAhoPJSFL = 280602
    gCapRepMensCapSaldCtaAhoPJCFL = 280603
    gCapRepMensCapSaldCtaAhoCMAC = 280604
    gCapRepMensCapSaldCtaAhoCRAC = 280605
    gCapRepMensCapSaldCtaPFPN = 280606
    gCapRepMensCapSaldCtaPFPJSFL = 280607
    gCapRepMensCapSaldCtaPFPJCFL = 280608
    gCapRepMensCapSaldCtaPFCMAC = 280609
    gCapRepMensCapSaldCtaPFCRAC = 280610
    gCapRepMensCapSaldCtaCTS = 280611
    gCapRepMensCapListGralCtas = 280612
    gCapRepMensCapSaldCtaCTSConvenio = 280614
    gCapRepMensCapSaldCtaCTSExternos = 280615
    gCapRepMensCapSaldCtaAhoFoncodes = 280616
    gCapRepMensCapSaldCtaAhoCooperativa = 280617
    gCapRepMensCapSaldCtaAhoEdipyme = 280618
    gCapRepMensCapSaldCtaPFFoncodes = 280619
    gCapRepMensCapSaldCtaPFCooperativa = 280620
    gCapRepMensCapSaldCtaPFEdpyme = 280621
    gCapRepMensBloqCapCTS = 280622
    gCapRepMensBloqCapCTSConvenio = 280623
    gCapRepMensBloqCapCTSExternos = 280624
    gCapRepMensCapSaldCtaCTSNascaPalpa = 280625
    gCapRepMensCapSaldCtaCTSCaneteMala = 280626
    gCapRepMensCapSaldCtaPFPNPJSFL = 280627
    gCapRepMensCapSaldCtaPFNascaPalpa = 280628
    gCapRepMensCapSaldCtaPFCaneteMala = 280629
    gCapRepMensBloqCapPF = 280630
    gCapRepMensCapSaldCtaAhoFinancieras = 280643 '***Agregado por ELRO el 20111105, según Acta 307-2011/TI-D
    gCapRepMensCapSaldCtaPFFinancieras = 280644 '***Agregado por ELRO el 20111105, según Acta 307-2011/TI-D
    
    'By Capi 07082008
    gCapRepExtractoCtasOP = 280640
    gCapRepExtractoCtasCTS = 280641
    gCapRepPagosWesternUnion = 280305
    '
    'By Capi 01102008
    gCapRepDiaPersoneriaPF = 280642
    
    'Promotores de Captaciones
    gCapRepVariosCapPromAho = 280709
    gCapRepVariosCapPromPF = 280710
    gCapRepVariosCapPromCTS = 280711
    gCapRepVariosCapPromotores = 280723
    
    'JUEZ 20140523 ***********************
    'Seguros de Tarjeta de Débito
    gCapExtCargoAfilSegTarjeta = 290001
    gCapExtCargoAfilSegTarjetaAho = 290002
    gCapExtCargoAfilSegTarjetaCTS = 290003
    'END JUEZ ****************************
    
    'Operaciones de Servicios
    gServCobSedalib = 300101
    gServCobHidrandina = 300102
    gServCobFideicomiso = 300103
    gServCobEdelnor = 300104
    gServCobSATTReciboDerecho = 300105
    gServCobSATTInfraccion = 300106
    gServCobSATTReciboDerechoOficEsp = 300107
    gServCobFoncodes = 300108 'Cobro de Foncodes
    gServCobPlanBici = 300109 'Cobro de Plan Bici
    gServCobServConv = 300110 'MADM 20110318
    gServCobAfilSegTarj = 300130 'JUEZ 20150112
    gServCobDebitoAuto = 300140 'JUEZ 20150130
    gServCobDebitoAutoEdit = 300141 'JUEZ 20150130
    gServCobSepelioPrima = 300150 'RECO 20151124
    gServActSepelioManual = 300151 'RECO 20151124
    gServCobPriSegSoat = 300155 'CTI2 FERIMORO : ERS034-2019    28082020
    
    'Otras Operaciones
    gOtrOpePagoFaltante = 300200
    gOtrOpeDepCtaBcoEfec = 300300
    
    'Ingresos
    gOtrOpeAhoOtrosIngresos = 300401
    gOtrOpeTransBancos = 300402
    gOtrOpeIngresosoCajaGeneral = 300403
    gOtrOpeComisionCMACLlam = 300404
    gOtrOpeTransferenciasEfect = 300405
    gOtrOpePremioCancPF = 300427
    gOtrOpeComisionRecaudoFocmacm = 300438 'FRHU 20150520 ERS022-2015
    gOtrOpePagoRecaudoCajeroCorresponsal = 300442 'RIRO20160110
    gOtrOpeRedondeoITF = 300499
    
    'Egresos
    gOtrOpeAhoOtrosEgresos = 300501
    gOtrOpeDepositoBancos = 300502
    gOtrOpeDevolucionCredPersonal = 300503
    gOtrOpeCastigoDevolucionCredPersonal = 300508
    gOtrOpeEgresoComisionDepCtaRecaudadora = 300523 'EJVG20120416
    '***Agregado por ELRO el 20120412, según OYP-RFC005-2012
    gOtrOpeDesParGas = 300524 'Desembolso A Rendir Para Otros Gastos
    gOtrOpeDesParVia = 300525 'Desembolso A Rendir Para Viaticos
    gOtrOpeDesParCaj = 300526 'Desembolso A Rendir Para Caja Chica
    '***Fin Agregado por ELRO*******************************
    gOtrOpeEgresoDevSobranteOtrasOpeChq = 300528 'EJVG20140213
    gOtrOpeEgresoDevSobranteOtrasOpeVoucher = 300529 'RIRO20140530 ERS017
    gOtrOpeEgresosDesemCajaChica = 300591
    gOtrOpeEgresosDirectos = 300527 ' RIRO 20130702 SEGUN TI-ERS083-2013
    gotrOpeDepUtilidadesEfect = 300532 'RIRO 20150701 ERS162 2015, Operacion Efectivo
    gotropeDepUtilidadesTrans = 300533 'RIRO 20150701 ERS162 2015, Operacion con Transferencia
    
    'Tarifas
    gOtrOpeDuplicadoTarjeta = 300601
    gOtrOpeVentaListados = 300602
    gOtrOpeConsatnciaCancelacionCredito = 300603
    gOtrOpeElaboracionContrato = 300604
    
    gComisionEmisionChequeEfectivo = 300630 'CTI4 ERS0112020
    gComisionEmisionChequeCargoCta = 300631 'CTI4 ERS0112020
    
    gOtrOpePenalidadEcoTaxi = 300800 'EJVG20120622
    gOtrOpe
    
    'Ingresos por Comisiones
    'JUEZ 20130411 *********************
    gComisionReprogCredito = 300901
    gComisionEvalPolEndosada = 300902
    gComisionDupTasacion = 300903
    gComisionConsultaRENIEC = 300904
    'END JUEZ **************************
    gComisionEnvioEstadoCta = 300905 'JUEZ 20130528
    'JUEZ 20130831 *********************
    gComisionDiversasAho = 300906
    'gComisionDiversasAhoGasto = 300907
    'gComisionDiversasAhoCom = 300908
    'END JUEZ **************************
    
    'COMISIONES
    'JUEZ 20150928 ***********************
    gComiAhoReposicionTarjeta = 321001 'Comisión Reposición de Tarjeta
    gComiAhoContMoneda = 321003 'Comisión Conteo Monedas
    gComiAhoDiversas = 321099 'Comisiones Diversas de Ahorros
    
    gComiCredEmisionRenovCF = 322001 'Comisión por Emisión y Renovación Carta Fianza Efectivo
    gComiCredModifCF = 322002 'Comisión por Modificación Carta Fianza
    gComiCredReprogCred = 322003 'Reprogramación de Créditos
    gComiCredEvalPolizaEnd = 322004 'Evaluación de Póliza Endosada
    gComiCredEstadoCta = 322005 'Envío de Estado de Cuenta
    gComiCredDupCronograma = 322006  'Duplicado Cronograma de Pago
    gComiCredConstNoAdeudo = 322007  'Constancia de no Adeudo
    gComiCredConstNoAdeudoCargoCta = 322008 'NDX ERS0112020 Constancia de no Adeudo con Cargo a Cuenta
    gCargoCtaComiCredConstNoAdeudo = 203061 'NDX ERS0112020 Comisiones Diversas de Créditos con Cargo a Cuenta
    gComiCredDiversasCargoCta = 322098 'NDX ERS0112020 Comisiones Diversas de Créditos con Cargo a Cuenta
    gComiCredDiversas = 322099 'Comisiones Diversas de Créditos
        
    gComiOtrServBusqRegSUNARP = 323001 'Servicio de Búsqueda de registro en SUNARP
    
    gComiDiversasAhoCredGasto = 300907
    gComiDiversasAhoCredCom = 300908
    'END JUEZ ****************************
    gComiDiversasAhoCredComVoucher = 300909  'CTI7 'OPEv2
    'EXTORNOS COMISIONES
    'JUEZ 20150928 ***********************
    gExtComiAhoDiversas = 371099 'Extorno Comisiones Diversas de Ahorros

    gExtComiCredEmisionRenovCF = 372001 'Extorno Comisión por Emisión y Renovación Carta Fianza Efectivo
    gExtComiCredReprogCred = 372003 'Extorno Reprogramación de Créditos
    gExtComiCredEvalPolizaEnd = 372004 'Extorno Evaluación de Póliza Endosada
    gExtComiCredEstadoCta = 372005 'Extorno Envío de Estado de Cuenta
    gExtComiCredDiversas = 372099 'Extorno Comisiones Diversas de Créditos
    
    gExtComiOtrServBusqRegSUNARP = 373001 'Servicio de Búsqueda de registro en SUNARP
    'END JUEZ ****************************
    
    'Operaciones de Cargo a cuenta
    gCompraMECargoCtaAhorro = 900027 'CTI6 ERS0112020
    gVentaMECargoCtaAhorro = 900028 'CTI6 ERS0112020
    gRecuadoCargoCtaAhorro = 900029 'CTI6 ERS0112020
    
    'Operaciones deVoucher
     gCompraMEVoucher = 900045 'CTI7 OpeV2
     gVentaMEVoucher = 900046 'CTI7 OpeV2
     gServicioRecaudoVoucher = 900047 'CTI7 OpeV2

    'Operaciones con Cheques
    gChqOpeRegistro = 900031
    gChqOpeModFechaValor = 900032
    gChqOpeMantenimiento = 900032 'EJVG20140124
    gChqOpeValorInmediata = 900033
    gChqOpeConsultaEstado = 900034
    gChqOpeDeposito = 900034 'EJVG20140124
    gChqOpeMantGirador = 900035 'MIOL 20130518 SEGUN RQ13521
    
    '***Agregado por ELRO el 20120627, según OYP-RFC024-2012
    'Operaciones con Vouchers
    gVouOpeRegistro = 920001
    gVouOpeEditar = 920002
    gVouOpeRegEnFec = 920003
    gVouOpeRegDif = 920004
    '***Fin Agregado por ELRO*******************************
    '***Agregado por ELRO el 20130711, según RFC1306270002****
    gCapConSerPagDeb = 301000 '***Servicio de Pago Debito
    gCapConSerPagDebCom = 301001 '***Servicio de Pago Debito Comision
    gCapConSerPagDebPen = 301002 '***Servicio de Pago Debito Penalidad
    gCapExtConSerPagDeb = 350206 '***Extorno Servicio de Pago Debito
    gCapExtConSerPagDebCom = 350207 '***Extorno Servicio de Pago Debito Comision
    gCapExtConSerPagDebPen = 350208 '***Extorno Servicio de Pago Debito Penalidad
    '***Fin Agregado por ELRO el 20130711, según RFC1306270002
    'FRHU 20141201 ERS048-2014
    gCapNotaDeCargo = 302001
    gCapNotaDeAbono = 302002
    'Extorno de Notas
    gCapExtNotaDeCargo = 330001
    gCapExtNotaDeAbono = 330002
    'FIN FRHU 20141201
    'Extornos Operaciones con Cheques
    gChqOpeExtRegistro = 900037
    gChqOpeExtModFechaValor = 900038
    gChqOpeExtDeposito = 900038 'EJVG20140312
    gChqOpeExtValorInmediata = 900039
    
    'Extorno de OtrasOperaciones
    gOtrOpeExtorno = 390000
    gCajaGenDesemPendienteExtorno = 350201
    
    'Giros
    gServGiroApertEfec = 310101
    gServGiroApertCargoCta = 310102
    gServGiroApertVoucher = 310104
    gServGiroComiEfec = 310201
    gServGiroComiCargoCta = 310202
    gServGiroComiEfecOA = 310203
    gServGiroCancEfec = 310301
    gServGiroCancCargoCta = 310302
    gServGiroCambDestin = 310401 'RECO20140512 ERS008-2014
    gServGiroAnulacion = 310402 'RECO20140512 ERS008-2014
    gServGiroComiAnul = 310403 'RECO20140512 ERS008-2014
    gServGiroComiMov = 310501 'RECO20140512 ERS008-2014
    gITFGiroCancelEfect = 990127 'NAGL 20190306 Según Correo
    
    'Extornos Servicios
    gServExtCobSedalib = 350101
    gServExtCobHidrandina = 350102
    gServExtCobFideicomiso = 350103
    gServExtCobEdelnor = 350104
    gServExtCobSATTReciboDerecho = 350105
    gServExtCobSATTInfraccion = 350106
    gServExtCobSATTReciboDerechoOficEsp = 350107
    gServExtCobFoncodes = 350108
    gServExtCobPlanBici = 350109
    
    gOtrOpeExtPagoFaltante = 350200
    '***Agregado por ELRO el 20120417, según OYP-RFC005-2012
    gOtrOpeExtDesParGas = 350202 'Extorno Desembolso A Rendir Para Otros Gastos
    gOtrOpeExtDesParVia = 350203 'Extorno Desembolso A Rendir Para Viaticos
    gOtrOpeExtDesParCaj = 350204 'Extorno Desembolso A Rendir Para Caja Chica
    '***Fin Agregado por ELRO*******************************
    gOtrOpeExtPenalidadEcoTaxi = 350205 'EJVG20120630
    gOtrOpeExtDepCtaBcoEfec = 350300
    
    
    gServExtGiroApertEfec = 360101
    gServExtGiroApertCargoCta = 360102
    gServExtGiroCancEfec = 360201
    gServExtGiroCancCargoCta = 360202


    gAhoOperacionesPendientes = 200920

    gAhoSolAprobOperacion = 280715
    gAhoAutorizaOperacion = 280716
    gAhoRechazaOperacion = 280717
    gAhoAnulaOperacion = 280718
    
    gAhoExtornoElectronicoCargo = 950001 'CTI7 OpeV2
    gAhoExtornoElectronicoDeposito = 950002 'CTI7 OpeV2
        
    gITFCobroCargo = 990101
    gITFCobroEfectivo = 990102
    gITFCobroEfectivoCred = 990103
    gITFCobroEfectivoPrend = 990104
    gITFCobroEfectivoAsumidoPF = 990109
    gITFCobroCargoPigno = 990110 ''***PEAC 20090515
    '***Agregado por ELRO el 20120206, según Acta 245-2011/TI-D
    gITFCobroCargoPF = 990111
    '***Fin Agregado por ELRO**********************************
    gITFCobroCargoArFi = 990112 'ALPA 20120420
    gITFCobroCargoEcoTaxi = 990113 'EJVG20120612
    gITFCobroPagoCuotaEcoTaxi = 990114 'EJVG20120612
    gITFCobroCargoArFiC = 990116 'ALPA 20120620
    gITFCobroCheque = 990123 'FRHU 20150528 ERS022-2015
    gITFCobroVoucher = 990124 'FRHU 20150528 ERS022-2015
    gITFCobroCMACCargo = 990301
    gITFCobroCMAC = 990302
    gITFCobroCMACCred = 990303
    gITFCobroCMACPrend = 990304

    gITFCobroCargoExt = 990201
    
    gITFCobroOpePignoVoucher = 990400 'CTI4 ERS0112020
    gITFCobroOpePignoCargoACta = 990401 'CTI4 ERS0112020
    
    '*** PEAC 20080616
    gCapRepSeguiCtasAhoPandero = 280239
    
    'by capi 20012009 constantes para modulo auditoria

    gCapMantParametros = 290100
    gCapConsParametros = 290200
    gAhoMantTasaInteres = 290300
    gAhoConsTasaInteres = 290400
    gPFMantTasaInteres = 290500
    gPFConsTasaInteres = 290600
    gPFCambioTasaInteres = 290700
    gPFPactoTasaInteres = 290800
    gCTSMantTasaInteres = 290900
    gCTSConsTasaInteres = 291000
    gCTSCambioTasaInteres = 291010 'JUEZ 20140304
    gAhoMantCuenta = 291100
    gPFMantCuenta = 291200
    gCTSMantCuenta = 291300
    gAhoBloqDesbCuenta = 291400
    gPFBloqDesbCuenta = 291500
    gCTSBloqDesbCuenta = 291600
    gCapSimulacionPF = 291700
    gCapRegistroTarjMag = 291800
    gCapBloqDesbTarjMag = 291900
    gCapCambClavTarjMag = 292000
    gCapMantBeneficiarios = 292100
    gCapConsBeneficiarios = 292200
    gAhoSolicitudOP = 292300
    gAhoEmisionOrdPago = 292400
    gAhoCertificacionOrdPago = 292500
    gAhoAnulacionOrdPago = 292600
    gAhoConsultaOrdPago = 292700
    gAhoRegTalonarioOrdPago = 292800
    gCapGeneraACL = 292900
    gColGeneraACL = 293000
    gAhoParamPeroneria = 293100
    gPFParamPeroneria = 293200
    gCTSParamPeroneria = 293300
    gCapNivelesAutoriz = 293400
    gCapNivGruposAutoriz = 293500
    gCapAprobRechAutoriz = 293600
    gCapSolicitudTasasPreferen = 293700
    gCapAprobRechTasasPreferen = 293800
    gCapExtornoTasasPreferen = 293900
    gCapConsultaTasasPreferen = 294000
    gCapExoneraITF = 294100
    gCapExoneraDsctoInactivas = 294200
    gCapBloqDesbParciales = 294300
    gPFDesbloqueo = 294400
    gCapAprobarTCEspecial = 294500
    gCapRegistroCampanaPremio = 294600
    gCapAsignaPremio = 294700
    gAhoRenovPlazoPanderito = 294800
    gCaptacCampanaReg = 294900 'JUEZ 20160418
    gCaptacCampanaEdit = 294901 'JUEZ 20160418
    gAhoMantTasaInteresCamp = 295000 'JUEZ 20160418
    gPFMantTasaInteresCamp = 295100 'JUEZ 20160418
    gCTSMantTasaInteresCamp = 295200 'JUEZ 20160418

'End by capi 20012009

    gAhoDepositoEnLoteEfec = 200257
    gAhoDepositoEnLoteCheq = 200258
    gAhoDepositoEnLoteCarg = 200259
    
    gAhoExtCargoCtaCompraME = 230375 'CTI6 ERS0112020
    gAhoExtCargoCtaVentaME = 230376    'CTI6 ERS0112020
    gAhoExtCargoServicioRecaudo = 230377    'CTI6 ERS0112020
    
End Enum
'VAPA20170612
Public Enum CCEOperacionOrig
        gCCETraOrdCtaAho = 930026
        gCCEComisonTraCCE = 930028
End Enum
'END VAPA 20170312


'PASI20160711 CCE
Public Enum CMACMCCEOperacionOrig
     gCMCCETransfOrdinaria = 930001
    gCMCCETransfPagoHaberes = 930002
    gCMCCETransfPagoProveedor = 930003
    gCMCCETransfPagoCTS = 930004
    gCMCCETransfPagoTarjCred = 930005
    gCMCCETransfOrdCargoCtaAho = 930006
    gCMCCETransfOrdCargoCtaCts = 930007
    gCMCCEComisionTransferenciaCCE = 930008 'VAPA20170119CCE
    gCMCCEComisionTransferenciaAho = 930100
    gCMCCEComisionTransferenciaCTS = 930101
    gCMCCEITFTransferenciaAho = 990125
    gCMCCETransfExtorno = 940001
End Enum
Public Enum CMACMCCEOperacionRec
    gCMCCEDepTransfOrdinaria = 930031
    gCMCCEDepTransfPagoHaberes = 930032
    gCMCCEDepTransfPagoProveedor = 930033
    gCMCCEDepTransfPagoCTS = 930034
    gCMCCEDepTransfITF = 990126
    gCMCCEDepComisionTransferencia = 930102
End Enum
Public Enum CMACMCCEConceptosOrig
    gCMCCEConceptoMontoEnvioTrCCE = 5300
    gCMCCEConceptoComisEnvioTrCCE = 5301
    gCMCCEConceptoComisionTrCMACM = 5302
    gCMCCEConceptoITFTrCMACM = 5303
End Enum
Public Enum CMACMCCEConceptosRec
    gCMCCECptoComisTrCMACMRec = 5304
    gCMCCECptoITFTrCMACMRec = 5305
End Enum
Public Enum CCEOperacion
    gCCETransfOrdinaria = 220
    gCCETransfPagoHaberes = 221
    gCCETransfPagoProveedor = 222
    gCCETransfPagoCTS = 223
    gCCETransfPagoTarjCred = 225
End Enum
Public Enum CCEInstrumento
    gCCETransferencia = 1
    gCCECheque = 2
    gCCEChequeArchivo = 4
    gCCESaldoTransferencia = 7
End Enum
Public Enum CCETRASesion
    TRAPresentados = 1
    TRADevoluciones = 2
    TRAAbono = 3 'VAPA20170328 CAMBIO 7 POR 3
End Enum
'VAPA20161019CCE
Public Enum CCECHESesion
    CHEPresentados = 1
    CHERechazados = 2
End Enum
Public Enum CCEEstados
    gCCEEstadoPreReg = 10
    gCCEEstadoConfAbo = 14
    gCCEEstadoDev = 13
End Enum
'end PASI

'By MAVM 23012009
Public Enum CodigosIdentificacionPistas
    gNumeroMovimiento = 1
    gCodigoCuenta = 2
    gCodigoPersona = 3
    gCodigoFuenteIngreso = 4
    gCodigoGarantia = 5
    gNumeroProcesoSeleccion = 6
    gCodigoPoliza = 7
    'By capi 25012009
    gNumeroSolicitud = 8
    gNumeroAprobacion = 9
    'ALPA 20090126****************************
    gNumeroContrato = 10
    gCodigo = 11
    gCodLineaCred = 12
    '*****************************************
End Enum

Public Enum TiposAccionesPistas
    gInsertar = 1
    gModificar = 2
    gEliminar = 3
    gAlertaError = 4
    gIngresarSistema = 5
    gSalirSistema = 6
    gConsultar = 7 'MAVM 20100726
End Enum

Public Enum CapEstAutorizaOpe 'Permisos Especiales, identificados como operaciones
    'Solicitud de aprobacion
    gAhoEstAprobOpeSolicitado = 1000
    gAhoEstAprobOpeAutorizado = 1001
    gAhoEstAprobOpeRechazado = 1002
    gAhoEstAprobOpeAnulado = 1003
End Enum

Public Enum CapPermisosEspeciales 'Permisos Especiales, identificados como operaciones
    gCapPermEspSaldCtaEmp = 905001
    gCapPermEspApertTasaPref = 905002
    gCapPermEspRetSinTarj = 905003
    gCapPermEspInicioDia = 905004
    gCapPermEspretSinFirmas = 905005
End Enum


Public Enum CaptacConcepto
    gConcCapital = 1
    gConcInteres = 2
    gConcProvision = 3
    gConcRevProvision = 4
    gConcProvisionNeg = 5
    
    gConcOpedeCompra = 6
    gConcOpedeVenta = 7
    gConcOpedeUtilidad = 8
    gConcOpedePerdida = 9
    gConcOpeRetOtraAge = 10
    gConcOpeRetxMaxOpe = 11
    
    gConcProvIntAdela = 12
    gConcRevProvIntAdela = 13
    
    gConcOpeDepOtraAge = 14
    
    gConcITFCliente = 20
    gConcITFAsumido = 21
    gConcRetSinTarj = 23
    
    '***Agregado por ELRO el 20120828, según OYP-RFC024-2012
    gConcPenTran = 24
    '***Fin Agregado por ELRO el 20120828*******************
    '***Agregado por ELRO el 20130716, según RFC1306270002****
    gComConSerPag = 25
    gPenConSerPag = 26
    '***Fin Agregado por ELRO el 20130716, según RFC1306270002
    
    
    '***************** Agregado por RIRO el 20130314 *****************
    
    cComiServRecaudoEmpresa = 15
    cComiServRecaudoCliente = 16
    
    ' FIN Agregado por RIRO el 20130314 ******************************
    'RECO20140512 ERS008-2014*****************************************
    gComServCambDest = 28
    gComServAnulGiro = 29
    gComServMovGiro = 30
    'RECO END********************************************************
    gConcComAsumidoEcotaxi = 31 'EJVG20140717
    gConcComEmisionCheque = 32 'CTI4 ERS0112020
End Enum

Public Enum CaptacMovTipo
    gCapMovApertura = 1
    gCapMovCancelAct = 2
    gCapMovCancelInact = 3
    gCapMovRetiroInac = 4
    gCapMovRetiroInt = 5
    gCapMovRetiro = 6
    gCapMovDeposito = 7
    gCapMovIntCap = 8
End Enum


Public Enum CaptacSubProdAhorros
    gCapSubProdAhoCorriente = 0
    gCapSubProdAhoNanito = 1
    gCapSubProdAhoPanderito = 2
    gCapSubProdAhoPandero = 3
    gCapSubProdAhoDestino = 4
End Enum

Public Enum CaptacSubProdPlazoFijo
    gCapSubProdPFPlazoFijo = 0
End Enum

Public Enum CaptacSubProdCTS
    gCapSubProdCTS = 0
End Enum

Public Enum CaptacInstServicios
    gCapServSedalib = 1
    gCapServHidrandina = 2
    gCapServTelefonica = 3
    gCapServFideicomiso = 4
    gCapServEdelnor = 5
    gCapServSATTRecDer = 6
    gCapServSATTInfraccion = 7
    gCapServSATTRecDerOficEsp = 8
End Enum

Public Enum CapServTipoComision
    gCapServTpoComMontoRecibo = 1
    gCapServTpoComPorcentaje = 2
    gCapServTpoMontoFijoMensual = 3
End Enum

Public Enum CaptacFlagServicios
    gCapServFlagRegistrado = 1
    gCapServFlagNoValido = 2
    gCapServFlagProcesado = 3
End Enum

Public Enum CaptacConvenios
    gCapConvUNT = 101
    gCapConvNarvaez = 102
    gCapConvSantaRosa = 103
    gCapConvJuanPabloII = 104
    gCapConvJuanPabloInst = 105
    gCapConvMarianoSantos = 106
End Enum

Public Enum CaptacConvTipoCuenta
    gCapConvTpoCtaPension = 1
    gCapConvTpoCtaGasto = 2
    gCapConvTpoCtaMora = 3
End Enum

Public Enum CaptacConvTipoReferencia
    gCapConvTpoRefDeducible = 1
    gCapConvTpoRefNoDeducible = 2
End Enum

Public Enum CapServUNTConcepto
    gUNTEsSaludPlan1 = 1
    gUNTEsSaludPlan4 = 2
    gUNTMatReg = 3
    gUNTMatExt = 4
    gUNTMatReg2daProf = 5
    gUNTMatExt2daProf = 6
    gUNTCuotaFacultad = 7
    gUNTCarnetUniv = 8
    gUNTCarnetBiblio = 9
    gUNTCurso2daMat = 10
    gUNTPrimRepAñoAcad = 11
    gUNTCurso3raMat = 12
    gUNTSegRepAñoAcad = 13
    gUNTCurso4taMat = 14
    gUNTTerRepAñoAcad = 15
    gUNTRecargPermExtra = 16
    gUNTCursoRepeticion = 17
    gUNTCursoCargo = 18
    gUNTExamenSuficiencia = 19
    gUNTInternet = 20
    gUNTFichaSocial = 21
    gUNTExamenAplazado = 22
    gUNTExamenRezagado = 23
    gUNTServicioMedico = 24
    gUNTEsSaludPlan3 = 25
End Enum

Public Enum CapTpoOperacionEstad
    gCapTpoOpeApertura = 1
    gCapTpoOpeAbono = 2
    gCapTpoOpeCargo = 3
    gCapTpoOpeCargoInact = 4
    gCapTpoOpeCancelacion = 5
    gCapTpoOpeCancelInact = 6
    gCapTpoOpeIntCapitalizado = 7
    gCapTpoOpeRetiroInteres = 8
End Enum

Public Enum GiroTarifaTipo
    gGiroTarTpoMonto = 10
    gGiroTarTpoPorcentaje = 11
End Enum

Public Enum GiroRelacPersona
    gGiroRelPersRemitente = 10
    gGiroRelPersDestinatario = 11
End Enum

Public Enum CuentasMovidas
    gCtasMovAhoApeEfec = 1001
    gCtasMovAhoApeCheque = 1002
    gCtasMovAhoApeTreansferencia = 1003
    gCtasMovAhoApeDesembolso = 1004
    gCtasMovAhoMovEfec = 1005
    gCtasMovAhoCancelacion = 1006
    gCtasMovAhoDepCheque = 1007
    gCtasMovAhoMovSinEfec = 1008
    
    gCtasMovPFApeEfec = 2001
    gCtasMovPFApeCheque = 2002
    gCtasMovPFApeTreansferencia = 2003
    gCtasMovPFMovEfec = 2004
    gCtasMovPFCancelacion = 2005
    gCtasMovPFMovSinEfec = 2006

    gCtasMovCTSApeEfec = 3001
    gCtasMovCTSApeCheque = 3002
    gCtasMovCTSApeTreansferencia = 3003
    gCtasMovCTSMovEfec = 3004
    gCtasMovCTSCancelacion = 3005
    gCtasMovCTSDepCheque = 3006
    gCtasMovCTSMovSinEfec = 3007
End Enum

Public Enum CapOrdPagTalEstado
    gCapTalOrdPagEstSolicitado = 1
    gCapTalOrdPagEstEnviado = 2
    gCapTalOrdPagEstRecepcionado = 3
    gCapTalOrdPagEstEntregado = 4
    gCapTalOrdPagEstExtornado = 5
End Enum


Public Enum CapNivRetCancEstado
    gCapNivRetCancEstSolicitado = 0
    gCapNivRetCancEstAprobado = 1
    gCapNivRetCancEstRechazado = 2
End Enum

Public Enum CapOrdPagTalTipo
    gCapTalOrdPagTpo25 = 1
    gCapTalOrdPagTpo50 = 2
    gCapTalOrdPagTpo100 = 3
End Enum

Public Enum CapFidConcepto
    gConcFidCapital = 7001
    gConcFidInteres = 7002
    gConcFidPortes = 7003
    gConcFidIntComp = 7004
    gConcFidMora = 7005
    gConcFidComVcdo = 7006
    gConcFidProtesto = 7007
    gConcFidGastos = 7008
End Enum


Public Enum ServConcepto
    gServConcImporte = 1
    gServConcComision = 2
    gServConcGastos = 3
    gServConcCostas = 4
    gServConcDerEmis = 5
End Enum

Public Enum CapTasaEspEstados
    gCapTasaEspSolicitada = 0
    gCapTasaEspAprobada = 1
    gCapTasaEspRechazada = 2
    gCapTasaEspAperturada = 3
    gCapTasaEspExtornada = 4
End Enum

Public Enum ColocEstado
    gColocEstSolic = 2000     'Colocaciones Solicitado
    gColocEstSug = 2001       'Colocaciones Sugerido
    gColocEstAprob = 2002     'Colocaciones Aprobado
    gColocEstRech = 2003      'Colocaciones Rechazado
    gColocEstDesemb = 2004    'ColoCaCiones Desembolso
    gColocEstVigNorm = 2020   'ColoCaCiones Vigente Normal
    gColocEstVigVenc = 2021   'ColoCaCiones Vigente VenCido
    gColocEstVigMor = 2022    'ColoCaCiones Vigente Moroso
    gColocEstRefNorm = 2030   'ColoCaCiones RefinanCiado Normal
    gColocEstRefVenc = 2031   'ColoCaCiones RefinanCiado VenCido
    gColocEstRefMor = 2032    'ColoCaCiones RefinanCiado Moroso
    gColocEstEmbarg = 2040    'ColoCaCiones Embargado
    gColocEstCancelado = 2050 'Colocaciones Cancelado
    gColocEstRefinanc = 2060  'Colocaciones Refinanciado
    gColocEstJudicial = 2070  'Colocaciones Paso a Judicial
    gColocEstRetirado = 2080  'Colocaciones Retirado
    gColocEstHonrada = 2090   'Colocaciones Honrada
    gColocEstDevuelta = 2091  'Colocaciones Devuelta
    gColocEstRenovada = 2092  'Colocaciones Renovada
    gColocEstTransferido = 2301  'Colocaciones Transferido 'FRHU 20150415 ERS022-2015
    gColocEstTransfCance = 2305  'Cancelado Transferido 'FRHU 20150415 ERS022-2015
    gColocEstModificada = 2093 'CF MODIFICADA 'WIOR 20120806
    '**
    gColPEstRegis = 2100     'Pignoraticio Registro
    gColPEstDesem = 2101     'Pignoraticio Desembolso
    gColPEstDifer = 2102     'Pignoraticio Desembolso
    gColPEstCance = 2103     'Pignoraticio Cancelado
    gColPEstVenci = 2104     'Pignoraticio Vencido
    gColPEstRemat = 2105     'Pignoraticio Rematado
    gColPEstPRema = 2106     'Pignoraticio Para Remate
    gColPEstRenov = 2107     'Pignoraticio Renovado
    gColPEstAdjud = 2108     'Pignoraticio Adjudicado
    gColPEstSubas = 2109     'Pignoraticio Subasta
    gColPEstAnula = 2110     'Pignoraticio Anulado
    gColPEstChafa = 2111     'Pignoraticio Chafaloneado
    gColPEstAnNoD = 2112     'Pignoraticio Anulado No Desembolsado
    '**
    gColocEstRecVigJud = 2201   'Recuperaciones Vigente Judicial
    gColocEstRecVigCast = 2202  'Recuperaciones Vigente castigad
    gColocEstRecCanJud = 2203   'Recuperaciones Cancelad Judicial
    gColocEstRecCanCast = 2204  'Recuperaciones Cancelad Castigad
    gColocEstRecRefiJud = 2205  'Recuperaciones Refinanciado Judicial 'FRHU 20150530 ERS022-2015
    
    '**** CMPL
    gPigEstRegis = 2801             'Pigno Registro
    gPigEstDesemb = 2802            'Pigno Desembolso
    gPigEstAmortiz = 2803           'Pigno Amortizacion
    gPigEstReusoLin = 2804          'Pigno Reuso de Linea
    gPigEstCancelPendRes = 2805     'Pigno Cancelado Pendiente de Rescate
    gPigEstRescate = 2806           'Pigno Rescatado
    gPigEstRemat = 2807             'Pigno Remate (Contratos activos)
    gPigEstRematPRes = 2808         'Pigno Remate (Pendientes de rescate)
    gPigEstRematPFact = 2809        'Pigno Rematado Pendiente de Facturacion
    gPigEstPResRematPFact = 2811    'Pigno Pendiente de Rescate Rematado Pendiente de Facturacion
    gPigEstRematFact = 2810         'Pigno Rematado y Facturado
    gPigEstAdjud = 2812             'Pigno Adjudicacion
    gPigEstAnula = 2813             'Pigno Anulado
    gPigEstRescateSobranteP = 2814  'Pigno Rescate sobrante de Pieza - CRSF
    
End Enum

Public Enum ColocRelacPers
    gColRelPersTitular = 20
    gColRelPersConyugue = 21
    gColRelPersCodeudor = 22
    gColRelPersRepresentante = 23
    gColRelPersRepresCodeudor = 24
    gColRelPersGarante = 25
    gColRelPersAnalista = 28
    gColRelPersApoderado = 29
    gColRelPersEstudioJuridico = 30
    gColRelPersAbogadoReponsable = 31
    gColRelPersJuzgado = 32
    gColRelPersJuez = 33
    gColRelPersSecretario = 34
    gColRelPersAcreedor = 35
    gColRelPersEncargado = 37
    gColRelPersAval = 38 'MADM 20111020
End Enum

Public Enum ColocRelacPersInfogas
    gColRelPersTitularCred = 10
    gColRelPersConsecionario = 11
    gColRelPersOperCertif = 12
    gColRelPersOperGarantia = 13
    gColRelPersNotario = 14
    gColRelPersAseguradora = 15
    gColRelTasacion = 16
    gColRelComisionEstructuracion = 17
End Enum

Public Enum ColocTipoTasa
    gColocTasaCompNormal = 200 'ColoCaCión Tasa de Interes Compensatorio Normal
    gColocTasaCompGraCia = 201 'ColoCaCión Tasa de Interes Compensatorio GraCia
    gColocTasaMoraNormal = 210 'ColoCaCión Tasa de Interes Moratorio Normal
    gColocTasaMoraGraCia = 211 'ColoCaCión Tasa de Interes Moratorio GraCia
End Enum

'Public Enum ColocGrupo
'    gColocCalendCod = 101
'    gColocCalendFlag = 102
'    gColocCalendApl = 103
'    gColocConcepto = 104
'    gColocConceptoApl = 105
'    gColocCalificacion = 106
'    gColocNota = 107
'End Enum

Public Enum ColocTipoCalend
    gColocCalendCodPFCF = 10    'Periodo Fijo - Cuota Fija
    gColocCalendCodPFCFPG = 11  'Periodo Fijo - Cuota Fija - Periodo de Gracia
    gColocCalendCodPFCC = 20   'Periodo Fijo - Cuota Creciente
    gColocCalendCodPFCCPG = 21  'Periodo Fijo - Cuota Creciente - Periodo de Gracia
    gColocCalendCodPFCD = 30  'Periodo Fijo - Cuota Decreciente"
    gColocCalendCodPFCDPG = 31   'Periodo Fijo - Cuota Decreciente - Periodo de Gracia
    gColocCalendCodFFCF = 40  'Fecha Fija - Cuota Fija
    gColocCalendCodFFCFPG = 41  'Fecha Fija - Cuota Fija - Periodo de Gracia
    gColocCalendCodFFCC = 50 'Fecha Fija - Cuota Creciente
    gColocCalendCodFFCCPG = 51   'Fecha Fija - Cuota Creciente - Periodo de Gracia
    gColocCalendCodFFCD = 60  'Fecha Fija - Cuota Decreciente
    gColocCalendCodFFCDPG = 61  ' Fecha Fija - Cuota Decreciente - Periodo de Gracia
    gColocCalendCodCL = 70   'Libre
End Enum

Public Enum ColocMiVivTramo
    gColocMiVivTramoNoConcesional = 1
    gColocMiVivTramoConcesional = 2
End Enum

Public Enum ColocCalendApl
    gColocCalendAplDesembolso = 0
    gColocCalendAplCuota = 1
    gColocCalendAplJudicial = 2
    gColocCalendAplOtros = 9
End Enum

Public Enum ColocCalendEstado
    gColocCalendEstadoPendiente = 0
    gColocCalendEstadoPagado = 1
End Enum

Public Enum ColocConcepto
    gColocConceptoCodCapital = 1000 ' Capital
    gColocConceptoCodInteresCompensatorio = 1100 'Interes Compensatorio
    gColocConceptoCodInteresMoratorio = 1101  'Interes Moratorio
    gColocConceptoCodInteresGracia = 1102  'Interes Gracia
    gColocConceptoCodInteresReprogramado = 1103 'Interes Re-programado
    gColocConceptoCodInteresSuspenso = 1104
    gColocConceptoCodInteresCompVencido = 1105
    gColocConceptoCodInteresCompDesagio = 1106
    gColocConceptoCodInteresDiferidoAmp = 1113  'WIOR 20150421 - Interes Diferido para Ampliados
    gColocConceptoCodGastoCentralRiesgos1 = 1202
    gColocConceptoCodGastoCentralRiesgos2 = 1203
    gColocConceptoCodGastoCentralRiesgos3 = 1204
    gColocConceptoCodGastoCentralRiesgos4 = 1205
    gColocConceptoCodGastoCentralRiesgos5 = 1206
    gColocConceptoCodGastoCentralRiesgos6 = 1207
    gColocConceptoCodGastoCentralRiesgos7 = 1208
    gColocConceptoCodGastoSeguro = 1210
    gColocConceptoCodGastoSeguro2 = 1211
    gColocConceptoCodGastoSeguro3 = 1212
    gColocConceptoCodGastoSeguro4 = 1213
    gColocConceptoCodGastoSeguro5 = 1214
    gColocConceptoCodGastoSeguro6 = 1215
    gColocConceptoCodGastoSeguro7 = 1217 '1218 'desgravemen 'ARCV
    gColocConceptoCodGastoBusquedaRegistral = 1216
    gColocConceptoCodGastoComisionFlat = 1218 '1217
    gColocConceptoCodGastoProtesto1 = 1219
    gColocConceptoCodGastoProtesto2 = 1220
    gColocConceptoCodGastoGastosCobranza = 1221
    gColocConceptoCodGastoTasacion = 1222
    gColocConceptoCodGastoComisionPreCancelacion = 1223 'JUEZ 20130925
    gColocConceptoCodGastoPenalidadPagoAdelantadoSoles = 1224
    gColocConceptoCodGastoPenalidadPagoAdelantadoDolares = 1225
    gColocConceptoCodGastoDsctoPlanillaSoles = 1227
    gColocConceptoCodGastoDsctoPlanillaDolares = 1228
    gColocConceptoCodGastoEstCta = 1232 'By Capi 27122007 gastos por estado de cuenta
    gColocConceptoCodGastoPolizaIncendioHipoteca = 1231 'peac 20080111 gastos por poliza contra incendio - hipoteca
    gColocConceptoCodGastoSuspenso = 1233
    gColocConceptoCodComisionEstructuracion = 1237
    gColocConceptoCodTasacionVehiculo = 1242
    gColocConceptoCodGastoPolizaVehiculoMultiriesgo = 1243 'BRGO 20111111
    gColocConceptoCodGastoIGV = 1244 'alpa 20120413
    gColocConceptoCodGastoVarios = 1299  'Gastos Varios
    gColocConceptoCodGastoMicroseguros = 1245 'WIOR 20120521
    gColocConceptoCodGastoSeguroMultiriesgo = 1246 'WIOR 20120521
    gColocConceptoCodGastoComisionReprogNat = 1247 'JUEZ 20130412
    gColocConceptoCodGastoComisionReprogJur = 1248 'JUEZ 20130412
    gColocConceptoCodGastoComisionEnvioEstCta = 1249 'JUEZ 20130529
    gColocConceptoCodGastoComisionMoraLeasing = 1250 'ALPA 20131109
    gColocConceptoCodGastoSeguroMultiriesgoMYPE = 1272 'RECO20150525 ERS023-2015
    gColocConceptoCodGastoPolizaIncendioHipotecaGracia = 1279 'LUCV20180601, Según ERS022-2018
    gColocConceptoCodCambCarteraNorm_Venc = 1401 ' Cambios de Cartera
    gColocConceptoCodCambCarteraVenc_Norm = 1402
    gColocConceptoCodCambCarteraMoro_JuCD = 1403
    gColocConceptoCodCambCarteraVenc_JuCD = 1404
    gColocConceptoCodCambCarteraNorm_JuCD = 1405
    gColocConceptoCodCambCarteraMoro_JuSD = 1406
    gColocConceptoCodCambCarteraVenc_JuSD = 1407
    gColocConceptoCodCambCarteraNorm_JuSD = 1408
    
    gColocConceptoCodCambCarteraMorR_JuSD = 1409 ' Cambios de Cartera Refinanciada
    gColocConceptoCodCambCarteraVenR_JuSD = 1410
    gColocConceptoCodCambCarteraNorR_JuSD = 1411
    gColocConceptoCodCambCarteraMorR_JuCD = 1412
    gColocConceptoCodCambCarteraVenR_JuCD = 1413
    gColocConceptoCodCambCarteraNorR_JuCD = 1414
    
    '*
    gColPConceptoCodCapital = 2000              'Pignoraticio: Capital
    gColPConceptoCodCapitalVencido = 2001       'Pignoraticio: Capital Vencido
    gColPConceptoCodInteresCompensatorio = 2100 'Pignoraticio: Interes Compensatorio
    gColPConceptoCodInteresMoratorio = 2101     'Pignoraticio: Interes Moratorio
    gColPConceptoCodInteresVencido = 2102     'Pignoraticio: Interes Compensatorio
    gColPConceptoCodTasacion = 2200             'Pignoraticio: Tasacion
    gColPConceptoCodCustodia = 2201             'Pignoraticio: Custodia
    gColPConceptoCodCustodiaVencida = 2202      'Pignoraticio: Custodia Vencida
    gColPConceptoCodCustodiaDiferida = 2203     'Pignoraticio: Custodia Diferia
    gColPConceptoCodImpuesto = 2204             'Pignoraticio: Impuesto
    gColPConceptoCodPreparaRemate = 2205        'Pignoraticio: PreparaRemate
    gColPConceptoCodComisionRemate = 2206       'Pignoraticio: Comision Remate
    gColPConceptoCodSobranteRemate = 2207       'Pignoraticio: Sobrante Remate
    
    '*** PEAC 20090313
    gColPConceptoCodSobranteAdjudicado = 2209       'Pignoraticio: Sobrante Adjudicado
    
    gColPConceptoCodCustodiaDiferidaImpuesto = 2208 'Pignoraticio: Sobrante Remate
    
    gColPConceptoCodCostoDuplicado = 2210       'Pignoraticio: Costo Duplicado
    
    gColPConceptoCodVtaNetaSubasta = 2211       'Pignoraticio: Venta Neta Subasta
    gColPConceptoCodImpuestoVtaSubasta = 2212   'Pignoraticio: Impuesto Venta Subasta
    gColPConceptoCodValorAdjudica = 2213        'Pignoraticio: Valor de Registro de Adjudicacion
    
    gColPConceptoCodOro14 = 2251                'Pignoraticio: Oro 14k
    gColPConceptoCodOro16 = 2252                'Pignoraticio: Oro 16k
    gColPConceptoCodOro18 = 2253                'Pignoraticio: Oro 18k
    gColPConceptoCodOro21 = 2254                'Pignoraticio: Oro 21k
    gColPConceptoCodMontoEntregar = 2271        'Pignoraticio: Monto Entregar
    
    gColPConceptoCodCambioCarteraNomalMoroso = 2301   'Pignoraticio: Cambio Cartera Nomal-Moroso
    gColPConceptoCodCambioCarteraMorosoNormal = 2302  'Pignoraticio: Cambio Cartera Moroso-Normal
    gColPConceptoCodCambioTasacionNormalMoroso = 2311 'Pignoraticio: Cambio Tasacion Normal-Moroso
    gColPConceptoCodCambioTasacionMorosoNormal = 2312 'Pignoraticio: Cambio Tasacion Moroso-Normal
    '*
    gColRecConceptoCodCapital = 3000                'Recuperaciones : Capital
    gColRecConceptoCodInteresCompensatorio = 3100   'Recuperaciones : Interes Compensatorio
    gColRecConceptoCodInteresMoratorio = 3101       'Recuperaciones : Interes Moratorio
    gColRecConceptoCodGastoCobranza = 3200          'Recuperaciones : Gastos Cobranza
    gColRecConceptoCodGastoComisionAbog = 3201      'Recuperaciones : Gastos Comision Abogado
    gColRecConceptoCodGastoAdmin = 3202             'Recuperaciones : Gastos Administrati Extra
    
    gColRecConceptoCodCastigGasto4 = 3501         'Recuperaciones : Castigar Gasto SK > Provision
    gColRecConceptoCodCastigIngreso6 = 3502       'Recuperaciones : Castigar Ingreso Provision > SK
    gColRecConceptoCodCastigCtaOrdenCap = 3503    'Recuperaciones : Castigar Cta Orden Capital
    gColRecConceptoCodCastigCtaOrdenIntCom = 3504 'Recuperaciones : Castigar Cta Orden Int Comp
    gColRecConceptoCodCastigCtaOrdenIntMor = 3505 'Recuperaciones : Castigar Cta Orden Int Mor
    gColRecConceptoCodCastigCtaOrdenGasto = 3506  'Recuperaciones : Castigar Cta Orden Gastos
    
    '*
    gColCFConceptoCodCapital = 4000                 'CF : Capital (Monto Carta Fianza)
    gColCFConceptoCodComision = 4100                'CF : Comision
    gColCFConceptoCodComisionMod = 4200                'CF : Comision por Modificacion AGREGADO POR WIOR 20120806
    'FRHU 20150520 ERS022-2015
    gColTransfConceptoCodCapital = 5000                'Transferido FOCMACM : Capital
    gColTransfConceptoCodInteresCompensatorio = 5100   'Transferido FOCMACM : Interes Compensatorio
    gColTransfConceptoCodInteresMoratorio = 5101       'Transferido FOCMACM : Interes Moratorio
    gColTransfConceptoCodGastoCobranza = 5200          'Transferido FOCMACM : Gastos Transferido
    'FIN FRHU 20150520
    
    '****   CMPL
    gColPigConceptoCodCapital = 8000                'Pigno : Capital
    gColPigConceptoCodInteresComp = 8100            'Pigno : Interes Compensatorio
    gcolPigConceptoCodInteresMora = 8101            'Pigno : Interes Moratorio
    gColPigConceptoCodGarantia = 8102               'Pigno : Valor de Tasacion
    gColPigConceptoCodTasacion = 8201               'Pigno : Tasacion
    gColPigConceptoCodComiServ = 8202               'Pigno : Comision de Servicio
    gColPigConceptoCodPenalidad = 8203              'Pigno : Penalidad por Pronto Pago
    gColPigConceptoCodCustodia = 8204               'Pigno : Custodia
    gColPigConceptoCodComiVencida = 8205            'Pigno : Comision Vencida
    gColPigConceptoCodPreparaRemate = 8206          'Pigno : Preparacion de Remate
    gColPigConceptoCodDuplicado = 8207              'Pigno : Costo Duplicado de Contrato
    gColPigConceptoCodCambioCartNormalVenc = 8301   'Cambio de Cartera de Normal a Vencido
    gcolPigConceptoCodCambioCartVencNormal = 8302   'Cambio de Cartera de Vencido a Normal
    gColPigConceptoCodPrecioVenta = 8401            'Pigno : Valor de Venta
    gColPigConceptoCodIgv = 8402                    'Pigno : Valor de Igv de la Venta
    gColPigConceptoCodGarantCustAge = 8501          'Pigno : Garan en Custodia (Contratos nuevos)en Agencia
    gColPigConceptoCodGarantTrans = 8502            'Pigno : Garan en Transito (Contratos nuevos)
    gColPigConceptoCodGarantCustBoveda = 8503       'Pigno : Garan en Custodia en Boveda
    gColPigConceptoCodGarantPendRescBoveda = 8504   'Pigno : Garan Pend Rescate en Boveda
    gColPigConceptoCodGarantPendRescTrans = 8505    'Pigno : Garan Pend Rescate en Transito
    gColPigConceptoCodGarantPendRescAge = 8506      'Pigno : Garan Pend Rescate Agencia
    gColPigConceptoCodGarantBienAdjBoveda = 8507    'Pigno : Bienes Adjudicados en Boveda
    gColPigConceptoCodGarantBienAdjTrans = 8508     'Pigno : Bienes Adjudicados en Transito
    gColPigConceptoCodGarantBienAdjCustAge = 8509   'Pigno : Bienes Adjudicados Custodia en Agencia
    gColPigConceptoCodGarantBienRealBoveda = 8510   'Pigno : Bien Realizable Boveda
    gColPigConceptoCodGarantBienRealTrans = 8511    'Pigno : Bien Realizable en Transito
    gColPigConceptoCodGarantBienRealCustAge = 8512  'Pigno : Bien Realizable Custodia en Agencia
    gColPigConceptoCodComisionMartillero = 8601     'Pigno : Comsion del Martillero
    gColPigConceptoCodPendCompensacion = 8602       'Pigno : Pendiente de Compensacion
    
    gColocConceptoCodCapitalVAC = 1020 ' Capital VAC
    gColocConceptoCodAjusteCalendCOFIDE = 1030 'Calendario CODIFE
    gColocConceptoCodCapitalReprogxDesNat = 1040 'EJVG20121214
    gColocConceptoCodCapitalReprogRoyaAmarilla = 1050 'JUEZ 20131022
    gColocConceptoCodCapitalReprogZonaInundada = 1060 'PASI 20170405
    gColocConceptoCodPoliza = 1120 'ARCV 12-03-2007
    gColocConceptoCodMontoHonrado = 1260 'WIOR 20131226 MONTO HONRADO
    gColocConceptoCodComisionHonramiento = 1261 'WIOR 20131226 COMISION HONRADO
End Enum

Public Enum ColocConceptoApl
    gColocConceptoAplTodosDC = 1 'Todas los D y C
    gColocConceptoAplTodosD = 2 'Todos los D
    gColocConceptoAplDesembolso = 3 'Solo un D
    gColocConceptoAplTodosC = 4  'Todas las C
    gColocConceptoAplCuota = 5 'Solo una C
    gColocConceptoAplNingun = 0
End Enum

Public Enum ColocCalific
    gColocCalificacionNormal = 0  'Calificación Normal
    gColocCalificacionPotencial = 1 ' Calificación Potencial
    gColocCalificacionDeficiente = 2 'Calificación Deficiente
    gColocCalificacionDudoso = 3 ' Calificación Dudoso
    gColocCalificacionPerdida = 4 ' Calificación Perdida
End Enum

Public Enum ColocNota
    gColocNotaNormal = 0  'Nota Normal
End Enum


Enum ColocLineaCredTasas
    gColocLineaCredTasasIntCompNormal = 1
    gColocLineaCredTasasIntCompEspecial = 2
    gColocLineaCredTasasIntMoratNormal = 3
    gColocLineaCredTasasIntMoratEspecial = 4
    gColocLineaCredTasasIntGracia = 5
    gColocLineaCredTasasIntCompVencida = 6
End Enum
    
Enum ColocLineaCredPlazo
    gColocLineaCredPlazoCorto = 1
    gColocLineaCredPlazoLargo = 2
End Enum

Enum ColocCredTipo
    gColocCredTipoNormal = 1
    gColocCredTipoRefinanciado = 2
End Enum

Enum ColocCredCondicion
    gColocCredCondNormal = 1 'Normal
    gColocCredCondRecurrente = 2 'Recurrente
    gColocCredCondParalelo = 3 'Paralelo
    gColocCredCondRefinan = 4 'Refinanciamiento
    gColocCredCondAdicional = 7 'Adicional EJVG20130503
End Enum

Enum ColocDestino
    gColocDestinoCapTrab = 1
    gColocDestinoActFijo = 2
    gColocDestinoMixto = 3
    gColocDestinoConsumo = 4
'    gColocDestinoOtros = 5

    gColocDestinoNoIndicado = 0
    gColocDestinoOtros = 9
    gColocDestinoAdqInmueble = 10
    gColocDestinoCambEstructPasivo = 14 'EJVG 20160225 ERS002-2016
    gColocDestinoRefracionVivienda = 20 'EAAS 20180727 ERS054-2018
    gColocDestinoRemodelacionVivienda = 21 'EAAS 20180727 ERS054-2018
    gColocDestinoAmpliacionVivienda = 22 'EAAS 20180727 ERS054-2018
    gColocDestinoMejoramientoVivienda = 23 'EAAS 20180727 ERS054-2018
    gColocDestinoAguaSaneamiento = 26 'EAAS 20180727 ERS054-2018
End Enum

Enum ColocTiposGracia
    gColocTiposGraciaPrimeraCuota = 0
    gColocTiposGraciaUltimaCuota = 1
    gColocTiposGraciaProrateada = 2
    gColocTiposGraciaConfigurable = 3
    gColocTiposGraciaExonerada = 4
    gColocTiposGraciaEnCuotas = 5
    gColocTiposGraciaCapitalizada = 6
End Enum

Enum ColocTiposDesembolso
    gColocTiposDesembolsoTotal = 0
    gColocTiposDesembolsoParcial = 1
End Enum

Enum ColocCalendConceptoProc
    gColocCalendConceptoProcSugerido = 1
    gColocCalendConceptoProcAprobado = 2
    gColocCalendConceptoProcDesembolsado = 3
    gColocCalendConceptoProcReprogramado = 4
    gColocCalendConceptoProcReEstructurado = 5
    gColocCalendConceptoProcOtro = 6
End Enum

Enum ColocTipoPago
    gColocTipoPagoEfectivo = 1
    gColocTipoPagoCheque = 2
    gColocTipoPagoDacionPago = 3
    gColocTipoPagoCargoCta = 4
    gColocTipoPagoVoucher = 5
    gColocTipoPagoCargoCtaAuto = 6 'WIOR 20140620
    gColocTipoPagoCajeroCorresponsal = 7 'ALPA 20160323
    gColocTipoPagoDeposito = 8 'CTI7 OPEv2
End Enum
'TIPO PAGO COMPRA Y VENTA
'CTI6 ERS0112020
Enum CVTipoPago
    gCVTipoPagoEfectivo = 1
    gCVTipoPagoCargoCta = 4
    gCVTipoPagoVoucher = 5
End Enum
Enum ColocEstadoRefinanc
    gColocEstadoRefinancSolicitado = 1
    gColocEstadoRefinancSugerido = 2
    gColocEstadoRefinancAprobado = 3
End Enum

Enum ColocMetasCred
    gColocMetasCredPymes = 1
    gColocMetasCredDsctoPlanilla = 2
    gColocMetasCredPlazoFijo = 3
    gColocMetasCredCts = 4
    gColocMetasCredUsosDiversos = 5
    gColocMetasCredAgricolas = 6
End Enum

Enum ColocTipoMetas
    gColocTipoMetasMenSual = 1
    gColocTipoMetasAnual = 2
End Enum

Enum ColocMotivRechazo
    gColocMotivRechazoFaltaLiquidez = 1
    gColocMotivRechazoExesoTrabAnalista = 2
    gColocMotivRechazoNegVivFueraArea = 3
    gColocMotivRechazoInsufCapacPago = 4
    gColocMotivRechazoFaltaGarantias = 5
    gColocMotivRechazoFaltaExperClienteAct = 6
    gColocMotivRechazoIncapAtenderNeg = 7
    gColocMotivRechazoFaltaRequisitos = 8
    gColocMotivRechazoInformFalsa = 9
    gColocMotivRechazoNoViveEnDirIndicada = 10
    gColocMotivRechazoNegInsuf = 11
    gColocMotivRechazoInestabVivCliente = 12
    gColocMotivRechazoMalReferenPers = 13
    gColocMotivRechazoMoroso = 14
    gColocMotivRechazoMalosAntecCred = 15
    gColocMotivRechazoNoCumpleCond = 16
    gColocMotivRechazoFaltaVisita = 17
    gColocMotivRechazoClientePrefVencido = 18
    gColocMotivRechazoNoHistCred = 19
    gColocMotivRechazoRetiroSolic = 20
    gColocMotivRechazoNoRecogioCheque = 21
    gColocMotivRechazoNoAceptoCondComite = 22
    gColocMotivRechazoNoAceptaMontoMenor = 23
    gColocMotivRechazoFaltaDocumen = 24
    gColocMotivRechazoNoEncontrado = 25
    gColocMotivRechazoMantSistema = 26
    gColocMotivRechazoCambioSistema = 27
    gColocMotivRechazoErrorProces = 28
    gColocMotivRechazoCambioGarante = 29
    gColocMotivRechazoBoletasFalsas = 30
End Enum

Enum ColocGarantRec
    gColocGarantRecDacion = 1
End Enum

Enum ColocGarantRecEstado
    gColocGarantRecEstadoRegistrado = 1
    gColocGarantRecEstadoCancelado = 2
    gColocGarantRecEstadoAnulado = 3
End Enum


'Constante de Parametros
Enum ColocPParametros
    gColPParamTasaCustodia = 3001
    gColPParamTasaTasacion = 3002
    gColPParamTasaImpuesto = 3003
    gColPParamTasaCustodiaVencida = 3004
    gColPParamTasaPreparaRemate = 3005
    gColPParamTasaComisionRemate = 3006
    gColPParamTasaIGV = 3007

    gColPParamPorcentajePrestamo = 3011
    gColPParamMaxCostoTasacion = 3012
    gColPParamMinPesoOro = 3013
    gColPParamLim1MontoPrestamo = 3014
    gColPParamLim2MontoPrestamo = 3015
    gColPParamToleranciaMontoPrestamo = 3016

    gColPParamNroImpresionesContrato = 3017
    gColPParamMaximoDiasDesembolso = 3018
    gColPParamPrecioOro10 = 3021
    gColPParamPrecioOro12 = 3022
    gColPParamPrecioOro14 = 3023
    gColPParamPrecioOro16 = 3024
    gColPParamPrecioOro18 = 3025
    gColPParamPrecioOro21 = 3026

    gColPParamMaxNroRenovac = 3031
    gColPParamCostoDuplicadoContrato = 3032
    gColPParamMaxDiasCustodiaDiferida = 3033
    gColPParamPorcentajeCustodiaDiferida = 3034
    gColPParamFactorPrecioBaseRemate = 3041
    gColPParamDiasAtrasoParaRemate = 3042
    gColPParamDiasAtrasoCartaVenc = 3043
    gColPParamNroRematesParaAdjudic = 3045
    gColPParamCostoNotifAvisoVenci = 3058 'DAOR 20070117
    gColPParamComiConstNoAdeudo = 3125 'JUEZ 20151229
    gColPParamComiDupCronograma = 3126 'JUEZ 20151229
    'JUEZ 20130415 ***************************
    gColPParamComEvalPolEndosada = 3201
    gColPParamComDupTasacion = 3202
    gColPParamComConsultaRENIEC = 3203
    'END JUEZ ********************************
    gColPParamMontoMinPreCancPersJur = 3209 'JUEZ 20130925
    gColPParamPorcPreCancPersJur = 3210 'JUEZ 20130925
    
    '===== CMPL
    gPigParamPorPrestamoCteA1 = 8001        'Pigno : Porcentaje de Prestamo para Clientes A1
    gPigParamPorPrestamoCteA = 8002         'Pigno : Porcentaje de Prestamo para Clientes A
    gPigParamPorPrestamoCteB = 8003         'Pigno : Porcentaje de Prestamo para Clientes B
    gPigParamPorPrestamoCteB1 = 8004        'Pigno : Porcentaje de Prestamo para Clientes B1
    
    gPigParamPorPrestamoJoyaBue = 8010      'Pigno : Porcentaje de Prestamo para Joyas Buenas
    gPigParamPorPrestamoJoyaReg = 8011      'Pigno : Porcentaje de Prestamo para Joyas Regulares
    gPigParamPorPrestamoJoyaMal = 8012      'Pigno : Porcentaje de Prestamo para Joyas Buenas
    
    gPigParamPrestamoPlazoMin = 8015        'Pigno : Plazo Minimo para el Prestamo
    gPigParamPrestamoPlazoMax = 8016        'Pigno : Plazo Maximo para el Prestamo
    gPigParamAmortizaMin = 8017             'Pigno : Amortizacion Mínima de un contrato
    gPiParamDiasMinComision = 8020
    gPigParamDiasCambioCartera = 8022       'Pigno : Dias para cambio de cartera
    gPigParamDiasDesembolso = 8023          'Dias Mínimos para Desembolso
    gPigParamDiasAtrasoCartaVenc = 8024     'Dias Minimos de Vencido para emision de Cartas de Vencimiento
    gPigParamDiasAtrasoSelecRemate = 8025   'Dias De Atraso min para selecciond de contratos a Remate
    gPigParamPorcComisMartillero = 8026     'Porcentaje de Comision del Martillero
    gPigAgeBovedaVal = 8040                 'Agencia donde se ubica Boveda de Valores
    
    '**DAOR 20070330********
    gColPParamPorComiCorresponsaliaMN = 3076 'Porcentaje de Comisión de Cobro por Corresponsalia en MN
    gColPParamPorComiCorresponsaliaME = 3077 'Porcentaje de Comisión de Cobro por Corresponsalia en ME
    gColPParamMonMinCobroCorresponsaliaMN = 3078 'Monto Mínimo de Comisión de Cobro por Corresponsalia en MN
    gColPParamMonMaxCobroCorresponsaliaMN = 3079 'Monto Máximo de Comisión de Cobro por Corresponsalia en MN
    gColPParamMonMinCobroCorresponsaliaME = 3080 'Monto Mínimo de Comisión de Cobro por Corresponsalia en ME
    
    '***********************
    
End Enum

' Contantes de Estados de Proceso de Recuperación de Garantias
Enum ColocPRecGarEstado
    gColPRecGarEstNoIniciado = 0
    gColPRecGarEstIniciado = 1
    gColPRecGarEstTerminado = 2
End Enum


' Contantes de Estados Creditos en el Proceso de Venta
Enum ColocPRecGarVentaEstado
    gColPRecGarVentaEnProceso = 0
    gColPRecGarVentaVendido = 1
    gColPRecGarVentaNoVendido = 2
    gColPRecGarVentaSobranteCobrado = 4
    gColPRecGarAdjudicado = 5
End Enum

'** Contantes de Operaciones
Enum ColocPOperaciones
    gColPOpeRegContrato = 120100         'Col Pig : Reg. Contrato
    gColPOpeDesembolso = 120200          'Col Pig : Desembolso
    gColPOpeDesembolsoEFE = 120201       'Col Pig : Desembolso EFECT
    gColPOpeDesembolsoAboCta = 120202    'By capi 05032009 Acta 022-2009
    gColPOpeDesembolsoAboCtaOtraAge = 120203    'PEAC 20090515
    gColPOpeDesembolsoAboCtaNva = 120204    'PEAC 20090515
    gColPOpeAnula = 120300               'Col Pig : Anula Contrato
    gColPOpeAnulNoDesemb = 120400        'Col Pig : Anula Contrato No Desemb
    gColPOpeDevJoyasNoDesemb = 120500    'Col Pig : Devolucion de Joyas No Desembolsadas
    gColPOpeModifDescJoyas = 120600      'Col Pig : Modifica Descrip Joyas
    gColPOpeBloqueoJoyas = 120700        'Col Pig : Bloqueo de Joyas
    gColPOpeDesbloqueoJoyas = 120800     'Col Pig : DesBloqueo de Joyas
    gColPOpeRenovacEFE = 121100             'Col Pig : Renovacion Efectivo
    gColPOpeRenovNorEFE = 121101            'Col Pig : Renovacion Nor EFEC
    gColPOpeRenovMorEFE = 121102            'Col Pig : Renovacion Mor EFEC
    gColPOpeRenovacCHQ = 121110             'Col Pig :  Renovacion Cheque
    gColPOpeRenovNorCHQ = 121111            'Col Pig : Renovacion Nor CHEQUE
    gColPOpeRenovMorCHQ = 121112            'Col Pig : Renovacion Mor CHEQUE
    
    gColPOpeRenovacVoucher = 121123         'CTI4 ERS0112020 Col Pig :  Renovacion Voucher
    gColPOpeRenovNorVoucher = 121121        'CTI4 ERS0112020 Col Pig : Renovacion Nor VOUCHER
    gColPOpeRenovMorVoucher = 121122        'CTI4 ERS0112020 Col Pig : Renovacion Mor VOUCHER
    
    gColPOpeRenovacCargoCta = 121130         'CTI4 ERS0112020 Col Pig :  Renovacion Cargo a Cuenta
    gColPOpeRenovNorCargoCta = 121131        'CTI4 ERS0112020 Col Pig : Renovacion Nor Cargo a Cuenta
    gColPOpeRenovMorCargoCta = 121132        'CTI4 ERS0112020 Col Pig : Renovacion Mor Cargo a Cuenta
    
    gColPOpeCancelacEFE = 121200            'Col Pig : Cancelacion
    gColPOpeCancelNorEFE = 121201               'Col Pig : Cancelacion Nor EFEC
    gColPOpeCancelMorEFE = 121202               'Col Pig : Cancelacion Mor EFEC
    gColPOpeCancelacCHQ = 121210
    gColPOpeCancelNorCHQ = 121211               'Col Pig : Cancelacion Nor CHEQUE
    gColPOpeCancelMorCHQ = 121212               'Col Pig : Cancelacion Mor CHEQUE
    
    gColPOpeCancelVoucher = 121220                  'CTI4 ERS0112020 Col Pig : Cancelacion Voucher
    gColPOpeCancelNorVoucher = 121221               'CTI4 ERS0112020 Col Pig : Cancelacion Nor Voucher
    gColPOpeCancelMorVoucher = 121222               'CTI4 ERS0112020 Col Pig : Cancelacion Mor Voucher
    gColPOpeCancelCargoCta = 121230                 'CTI4 ERS0112020 Col Pig : Cancelacion Cargo a Cuenta
    gColPOpeCancelNorCargoCta = 121231              'CTI4 ERS0112020 Col Pig : Cancelacion Nor Cargo a Cuenta
    gColPOpeCancelMorCargoCta = 121232              'CTI4 ERS0112020 Col Pig : Cancelacion Mor Cargo a Cuenta
    
    gColPOpeImpDuplicado = 121600        'Col Pig : Imprimir Duplicado
    gColPOpeDevJoyas = 121800            'Col Pig : Devol Joyas
    gColPOpeCobCusDiferida = 121900      'Col Pig : Cobrar Custod Diferida
    gColPOpeVtaRemate = 122000           'Col Pig : Venta Joyas en Remate
    gColPOpeAboSobCta = 122100           'Col Pig : Abono Sobrante Remate
    gColPOpePagSobrante = 122200         'Col Pig : Pago de Sobrante Remate
        
    '*** PEAC 20090313
    gColPOpePagSobraAdjudicado = 122300 'Col Pig : Pago de Sobrante Adjudicado
    
    gColPOpeAdjudica = 122500            'Col Pig : Adjudica Joyas
    gColPOpeVentaLote = 122600            'DAOR 20080201, Col Pig : Venta en Lote de Cred. Adj.
    gColPOpeVtaSubasta = 122800          'Col Pig : Venta Joyas en Subasta

    gColPOpeAmortizEFE = 121000             'Col Pig : Ammortizacion Efectivo
    gColPOpeAmortNorEFE = 121001            'Col Pig : Amortizacion Nor EFEC
    gColPOpeAmortizCHQ = 121010             'Col Pig :  Amoritizacion Cheque
    gColPOpeAmortNorCHQ = 121011            'Col Pig : Amortizacion Nor CHEQUE

    '**** Operaciones Cambio de Cartera (Cierre Diario)
    gColPOpeCambCartera = 125800
    gColPOpeCambCartNorMor = 125801
    gColPOpeCambCartMorNor = 125802
    gColPOpeCambValTasac = 125900
    gColPOpeCambValTasVigVen = 125901
    gColPOpeCambValTasVenVig = 125901

    '**** Operaciones RECEPCION EN OTRA CMAC
    gColPOpeRenovacEnOtCjEFE = 126100           'Col Pig : Renovac EN Otra CMAC
    gColPOpeRenovNorEnOtCjEFE = 126101        'Col Pig : Renovac Normal EN Otra CMAC EFECTIVO
    gColPOpeRenovMorEnOtCjEFE = 126102        'Col Pig : Renovac Morosa EN Otra CMAC EFECTIVO
    gColPOpeRenovacEnOtCjCHQ = 126110
        gColPOpeRenovNorEnOtCjCHQ = 126111        'Col Pig : Renovac Normal EN Otra CMAC CHEQUE
        gColPOpeRenovMorEnOtCjCHQ = 126112        'Col Pig : Renovac Morosa EN Otra CMAC CHEQUE
    gColPOpeCancelacEnOtCjEFE = 126200           'Col Pig : Cancelacion Efectivo EN Otra CMAC
        gColPOpeCanceNorEnOtCjEFE = 126201        'Col Pig : Cancelac Normal EN Otra CMAC EFECTIVO
        gColPOpeCanceMorEnOtCjEFE = 126202        'Col Pig : Cancelac Morosa EN Otra CMAC EFECTIVO
    gColPOpeCancelacEnOtCjCHQ = 126210           'Col Pig : Cancelacion Efectivo EN Otra CMAC
        gColPOpeCanceNorEnOtCjCHQ = 126211        'Col Pig : Cancelac Normal EN Otra CMAC CHEQUE
        gColPOpeCanceMorEnOtCjCHQ = 126212        'Col Pig : Cancelac Morosa EN Otra CMAC CHEQUE
    gColPOpeAmortizEnOtCjEFE = 126300           'Col Pig : Amortiz EN Otra CMAC
        gColPOpeAmortNorEnOtCjEFE = 126301        'Col Pig : Amortiz Normal EN Otra CMAC EFECTIVO
    gColPOpeAmortizEnOtCjCHQ = 126310
        gColPOpeAmortNorEnOtCjCHQ = 126311        'Col Pig : Amortiz Normal EN Otra CMAC CHEQUE
    
    '**** Operaciones LLAMADA A OTRA CMAC (Se utiliza el de creditos)

    '**** Operaciones EXTORNOS
    gColPOpeExtRenov = 129100             'Col Pig : Extorno Renovacion
    gColPOpeExtDevJoyas = 129200          'Col Pig : Extorno Devol Joyas
    gColPOpeExtCustodDifer = 129300       'Col Pig : Extorno Custodia Difer
    gColPOpeExtCance = 129400             'Col Pig : Extorno Cancelacion
    gColPOpeExtPagoParc = 129401          'Col Pig : Extorno Pago ParciaL 'PEAC 20161212
    gColPOpeExtDesemb = 129500            'Col Pig : Extorno Desembolso
    gColPOpeExtDesembAmpliado = 129501    'Col Pig : Extorno Desembolso Ampliado 'RECO20140208 ERS002
    gColPOpeExtDupli = 129600             'Col Pig : Extorno Duplicado

    gColPOpeExtRenovEOCmac = 129801       'Col Pig : Extorno Renov En ot Cmac
    gColPOpeExtCanceEOCmac = 129802       'Col Pig : Extorno Cance En Ot Cmac
    
    'By Capi 01102008
    gColRepCredRefinan = 108351

    '****** Operaciones Pignoraticio - CMPL
    gPigOpeRegContrato = 150100         'Col Pig : Reg. Contrato
    gPigOpeDesembolso = 150200          'Col Pig : Desembolso
        gPigOpeDesembolsoEFE = 150201   'Col Pig : Desembolso EFECT
        gPigOpeDesembolsoUL = 150202    'Col Pig : Desembolso EFECT
    gPigOpeAnulacion = 150300           'Col Pig : Anula Contrato
    gPigOpeAnulNoDesemb = 151400        'Col Pig : Anula Contrato no Desembolsado
    gPigOpeDevJoyasNoDesemb = 151400    'Col Pig : Devolucion de Joyas No Desembolsadas
    gPigOpeModifDescJoyas = 150500      'Col Pig : Modifica Descrip Joyas
    gPigOpeReusoLinea = 150600          'Col Pig : Reuso de Linea
    gPigOpeBloqueoJoyas = 150700        'Col Pig : Bloqueo de Joyas
    gPigOpeDesbloqueoJoyas = 150800     'Col Pig : DesBloqueo de Joyas
    gPigOpeAmortizEFE = 150900          'Col Pig : Renovacion Efectivo
        gPigOpeAmortNorEFE = 150901     'Col Pig : Amortizacion Nor EFEC
        gPigOpeAmortMorEFE = 150902     'Col Pig : Amortizacion Mor EFEC
        gPigOpeAmortNorUL = 150903      'Col Pig : Amortizacion Nor Uso de Linea
        gPigOpeAmortMorUL = 150904      'Col Pig : Amortizacion Mor Uso de Linea
    gPigOpeCancelacEFE = 151000         'Col Pig : Cancelacion
        gPigOpeCancelNorEFE = 151001    'Col Pig : Cancelacion Nor EFEC
        gPigOpeCancelMorEFE = 151002    'Col Pig : Cancelacion Mor EFEC
    gPigOpeImpDuplicado = 151100        'Col Pig : Imprimir Duplicado
    gPigOpeDevJoyas = 151200            'Col Pig : Devol Joyas
    gPigOpeCobCusDiferida = 151300      'Col Pig : Cobrar Custod Diferida
    gPigOpeDespRemesa = 151501          'Col Pig : Despacho Remesa de Joyas
    gPigOpeRecepRemesa = 151502         'Col Pig : Recepcion Remesa de Joyas
    gPigOpeExtornos = 159000            'Col Pig : Extornos
    gPigOpeExtRescateJoya = 159100       'Col Pig : Extorno Rescate Joya
    gPigOpeExtDesembolso = 159200       'Col Pig : Extorno Desembolso
    gPigOpeExtAmortizEFE = 159300       'Col Pig : Extorno Renovacion Efectivo
    gPigOpeExtCancelacEFE = 159400      'Col Pig : Extorno Cancelacion
    gPigOpeExtReusoLinea = 159500       'Col Pig : Extorno Reuso de Linea
    gPigOpeExtCobCusDiferida = 159600   'Col Pig : Extorno Cobrar Custod Diferida
    gPigOpeAnulaVentaJoya = 159700      'ColPig  : Anula Venta de Joya en Tienda
    gPigOpeAnulaVentaJoyaRemate = 159701
    gPigOpeExtImpDuplicado = 159800     'Col Pig : Extorno Imprimir Duplicado
    gPigOpeExtPagoSobrante = 159900     'Col Pig : Extorno de Pago de Sobrante
    gPigOpeExtPagoSobranteP = 159901    'Col Pig : Extorno de Pago de Sobrante Pieza
    gPigOpeRegSobrantes = 154000        'ColPig  : Registro de Sobrantes
    gPigOpePagoSobrantes = 152100       'ColPig  : Pago de Sobrantes
    gPigOpeVentaJoya = 152800           'ColPig  : Venta de Joya en Tienda
    gPigOpeSobraPieza = 152900           'ColPig : Sobrante de Pieza
    gPigOpeVentaRemate = 152000         'ColPig  : Venta de Joyas en Remate
    gPigOpeDespContNuevo = 151501       'ColPig  : Despacho de Contratos Nuevos
    gPigOpeRecepContNuevo = 151502      'ColPig  : Recepcion de Contratos Nuevos
    gPigOpeDespPendResc = 151503        'ColPig  : Despacho de Contratos Pendiente de Rescate
    gPigOpeRecepPendResc = 151504       'ColPig  : Recepcion de Contratos Pendientes de Rescate
    gPigOpeDespPendRescMay = 151505     'ColPig  : Despacho de Contratos Pendientes de Rescate > a 30 dias
    gPigOpeRecepPendRescMay = 151506    'ColPig  : Recepcion de Contratos Pendientes de Rescate > a 30 dias
    gPigOpeDespRemateVenc = 151507      'ColPig  : Despacho Para Remate de Contratos Vencidos
    gPigOpeRecepRemateVenc = 151508     'ColPig  : Recepcion para Remate de Contratos Vencidos
    gPigOpeDespDevRemateVenc = 151509   'ColPig  : Despacho Dev Contratos Vencidos no Rematados
    gPigOpeRecepDevRemateVenc = 151510  'ColPig  : Recepcion Dev Contratos Vencidos no Rematados
    gPigOpeDespJoyasVenta = 151511      'ColPig  : Despacho de Joyas para Venta
    gPigOpeRecepJoyasVenta = 151512     'ColPig  : Recepcion de Joyas para Venta
    gPigOpeDespDevJoyasVenta = 151513   'ColPig  : Despacho de Joyas
    gPigOpeRecepDevJoyasVenta = 151514  'ColPig  : Recepcion de Joyas
    gPigOpeDespJoyasVentaRemate = 151515      'ColPig  : Despacho de Joyas para Ventas y Remates
    gPigOpeRecepJoyasVentaRemate = 151516     'ColPig  : Recepcion de Joyas para Ventas y Remates
    gPigOpeDespDevJoyasVentaRemate = 151517   'ColPig  : Despacho de Joyas de Ventas y Remates
    gPigOpeRecepDevJoyasVentaremate = 151518  'ColPig  : Recepcion de Joyas de Ventas y Remates
    
    gIngresarSalirSistema = 900100
    
    
    '*** PEAC 20090122 - VARIABLES PARA LAS PISTAS - PIGNORATICIO
    gPigRegistrarContrato = 190860
    gPigConsultarContro = 190870
    gPigAnularContrato = 190880
    gPigBloquearContrato = 190890
    gPigRescatarJoyas = 190900
    gPigPrepararSubasta = 190910
    gPigSubastarLote = 190920
    gPigPrepararVenta = 190930
    gPigVenderLote = 190940
    gPigPrepararAdjudicacion = 190950
    gPigAdudicarLote = 190960
    gPigPagoSobranteAdjudicado = 190970

    '*** PEAC 20090126 - VARIABLES PARA LAS PISTAS - CREDITOS
    gCredRegistrarActualizaSoliCred = 190260
    gCredRegistraActualizaGarantia = 190310
    gCredRegistraSugerenciaAnalista = 190390
    gCredRegistraAprobacionCredito = 190410
    gCredReprogramarCredito = 190470
    gCredRegistrarCartaFianza = 190010
    gCredRegistrarEvaluacionCred = 190391 'JUEZ 20120915
    gCredMantenimientoEvaluacionCred = 190392 'JUEZ 20120915
    gCredVerificacionEvaluacionCred = 190393 'JUEZ 20120915
    gCredExtornoVerificacionEvaluacionCred = 190394 'WIOR 20120917
    gCredConsultarEvaluacionCred = 190395 'LUCV20181220, Anexo01 de Acta 199-2018
    gCredEliminacionEvaluacionCred = 190396 'LUCV20181220, Anexo01 de Acta 199-2018
    gCredConsultarFteIngreso = 190397 'LUCV20181220, Anexo01 de Acta 199-2018
    
    gCredSugerenciaCF = 190030 'MAVM 20102106
    gCredAprobacionCF = 190040 'MAVM 20102106
    gCredRechazarCF = 190050 'MAVM 20102506
    gCredRetirarCFAprobada = 190060 'MAVM 20102506
    gCredDevolverCFEmitida = 190070 'MAVM 20102506
    gCredEmisionCF = 190080 'MAVM 20102506
    gCredHonrarCF = 190090 'MAVM 20102506
    gCredMantTarfCF = 190091 'MAVM 20102506
    gCredRenovacionCF = 190092 'MAVM 20102506
    gCredModificacionModCF = 190094 'WIOR 20130311
    
    gCredRegistrarGastosLote = 190520
    gCredExonerarPenalidadCancel = 190530
    gCredAdministrarGastos = 190540
    gCredRechazarCreditos = 190420
    gCredRetirarCreditos = 190430
    gCredRechazarSolicitud = 190465
    gCredRechazarSugerencia = 190460
    gCredRegistrarGarantiaCli = 190310
    gCredRegistrarGravamen = 190340
    gGarantiaBloqueoLegal = 190341
    gGarantiaVerificaLegal = 190342
    gGarantiaObviarDescobertura = 190343
    gCredLiberarBloquearGarantia = 190350
    gCredActualizarMetodoLiquid = 190500 'MAVM 20120328
    gCredConsultaHistorialCred = 190630 'MAVM 20100726
    gCredAdminPrepagosNormales = 190710
    gCredAdminPrepagosHipoteca = 190700
    gCredAutorizarAprobacion = 190760
    gCredLineaCred = 190770
    gCredPerdonarMora = 190510

    '*** PEAC 20090126 - VARIABLES PARA LAS PISTAS - RECUPERACIONES
    gRecRegistrarExpediente = 190970
    gRecRegistrarActuaProcesal = 190980
    gRecAsignarGastos = 190990
    gRecRegistrarMetodoLiquid = 191000
    
    gsPersSaldosConsol = 191330 'WIOR 20130122
    'FRHU 20150520 ERS022-2015
    gCredDistribucionPagoTransfFocmacm = 191340
    gCredCancelacionPagoTransfFocmacm = 191350
    'FIN FRHU 20150520
    'FRHU 20141115 MEMO-2766-2014
    gConsultarFirma = 900101
    gMantenimientoFirma = 900102
    gPersonaRegistro = 900110 'LUCV20181220, Anexo01 de Acta 199-2018
    gPersonaMantenimiento = 900111 'LUCV20181220, Anexo01 de Acta 199-2018
    gPersonaConsulta = 900112 'LUCV20181220, Anexo01 de Acta 199-2018
    'FIN FRHU
    
    '*** PEAC 20161111 - Pagos Parciales en Pignoraticio
    gColPOpePagoParcialEfectivo = 121400
        gColPOpePagoParcialNorEfectivo = 121401        'Col Pig : Pago parcial Normal efectivo
    gColPOpePagoParcialCheque = 121410
        gColPOpePagoParcialNorCheque = 121411        'Col Pig : Pago parcial Normal cheque
    gColPOpePagoParcialNorVoucher = 121421      'CTI4 ERS0112020
    gColPOpePagoParcialNorCargoCta = 121431     'CTI4 ERS0112020
    gColPOpeRecuperaContratoAdjVoucher = 122910 'CTI4 ERS0112020
    gColPOpeRecuperaContratoAdjCargoCta = 122920 'CTI4 ERS0112020
    
End Enum
    
'Reportes de Creditos
Enum ColCredReportes
    gColCredRepReportesCreditos = 108000 'Reportes de Creditos
    gColCredRepRepDiarios = 108100 'Reportes Diarios
    gColCredRepIngxPagoCred = 108101  'Ingresos por Pago de Créditos
    gColCredRepDesemEfect = 108102  'Desembolsos Efectuados
    gColCredRepSalCarVig = 108103 'Saldos de Cartera Vigente
    gColCredRepResSalCarxAna = 108104 'Resumen de Saldos de Cartera por Analista
    gColCredRepMoraInst = 108105 'Mora Institucional
    gColCredRepMoraxAna = 108106 'Mora por Analista
    gColCredRepCredProtes = 108107 'Créditos Protestados
    gColCredRepRenCarAnali = 108150 'Rentabilidad de creditos por analista
    gColCredRepCredSaldosDiarios = 108118 'Creditos Saldos diarios x Agencia y Fecha
    gColCredCanXNAtendidos = 108151 'Rentabilidad de creditos por analista
    
    
    'gColCredRepCredRetir = 108108 'Créditos Retirados
    gColCredRepCredRetir = 108190 'Créditos Retirados
    gColCredRepCredRecha = 108191 'Créditos Rechazados
    gColCredRepCredAnula = 108192 'Créditos Anulados
    
    gColCredRepCredCancel = 108109 'Créditos Cancelados
    gColCredRepListCuotaLibre = 108180 'Listados Cuota Libre
    gColCredRepAtraPagoCuotaLib = 108181 'Atraso en Pago con Cuota Libre
    gColCredRepCredVigconCuoLibre = 108182 'Creditos Vigentes con Cuota Libre
    
        
        
    gColCredRepMensuales = 108200 'Reportes Mensuales
    gColCredRepCredVigIntDeven = 108201 'Créditos Vigentes - Int.Devengado
    gColCredRepEstMensual = 108202 'Estadística Mensual
    gColCredRepCredVigArqueo = 108203 'Créditos Vigentes - Arqueo
    gColCredRepCredDesmMayores = 108204 'Creditos Desembolsados Mayores
    gColCredRepReportesVarios = 108300 'Reportes Varios
    gColCredRepProgPagosxCuota = 108301 'Programación de Pagos de Cuotas
    gColCredRepVisitaCobroCuotas = 108302 'Visita para cobro de Cuotas
    gColCredRepClientesNCuotasPend = 108303 'Clientes con N Cuotas Pendientes
    gColCredRepIngresosxGasto = 108304 'Reporte de Ingresos por Gastos
    gColCredRepCredxUbiGeo = 108305 'Créditos por Ubicación Geográfica
    gColCredRepDatosReqMora = 108306 'Datos Para requerimiento de mora
    gColCredRepDatosPromotor = 108307 'Datos Para Creditos por Promotor
    
    '** DAOR 20070313, Reportes de Corresponsalia*******************************
    gColCredRepCorresponsalia = 108330 'Reportes de Corresponsalia
    gColCredRepCreditosDesBcoBac = 108331 'Creditos desembolsados en agencias del banco de la nación
    '***************************************************************************
    'By capi 28122007 garantias inscritas
    gColCredRepGarantiasInscritas = 108386
    '
    gColCredRepCreditosAprExoReglamento = 108326 'DAOR 20070418, Reporte de Creditos Aprobados con Exoneración al Reglamento
    gColCredRepMontoDesembolsadoPorLineas = 108327 'DAOR 20070418, Reporte de Montos desembolsados por línea de crédito
    gColCredRepEstadosCuentaCredito = 108328 'DAOR 20070717, Estados de cuenta de crédito
    gColCredRepConCarxAnalista = 108120 'DAOR 20070717, Consolidado de cartera por analista
    gColCredRepActaComiteCredAprobados = 108121 'PEAC 20070822 REPORTES E ACTAS DE COMITE DE CREDITOS APROBADOS
    gColCredRepXTipoCondProd = 108122 'PEAC 20070822 REPORTES por tipo de condicion de producto
    gColCredRepDetalleCuotasXCobrar = 108123 'PEAC 20070903 reporte detalle de cuotas por cobrar
    gColCredRepCreditosXCampanha = 108124 'PEAC 20070822 REPORTES de credios por campañas
    gColCredRepConvCasillero = 108209 'PEAC 20070921 REPORTES de COBRANZA DIARIA (CASILLERO)
    gColCredRepLavadoDinero = 108340 'By Capi 28012008
    gColCredRepREULavadoDinero = 108341 'By Capi 28012008
    gColCredRepDUDLavadoDinero = 108342 'By Capi 30012008
    
    gColCredRepComVigEEFF = 108329 'PEAC 20071228 REPORTES de creditos comerciels vigentes con fecha eeff
    gColCredRepPoliIncendio = 108308 'PEAC 20071229 REPORTES de polizas contra insendio
    
    '*** PEAC 20080215 CONTROL DE SOLICITUDES
    gColCredRepDiasTranscDesdeSoli = 108141
    gColCredRepSoliRechazadas = 108142
    gColCredRepSoloEnEstadoSoli = 108143
    gColCredRepSoliProcesadas = 108144
    
    gColCredRepCredVencPaseCastigo = 108335 '*** PEAC 20080220
    gColCredRepCliConDistinTiposCred = 108336 '*** PEAC 20080221
    gColCredRepRiegCredOfiBN = 108337 '*** GITU 20080925
    gColCredRepAECIReporte01 = 108921 '*** PEAC 20080308
    gColCredRepClientesHisNegativo = 108350 '*** PEAC 20080804
    gColCredRepClientesPotencialesSinCredVig = 108360 '*** PEAC 20080923
    gColCredRepNumYSaldoCredPorProductoConsol = 108210 '*** PEAC 20080924
    
    '********************************************************************************
    gColRecRepCredJudToCastigar = 138037 'PEAC 20070912 REPORTES de credios judiciales para castigar
    gColRecRepGastosJudCas = 138038  'PEAC 20071010 REPORTES de gastos de creditos jud y cas
    gColRecRepCredSaldosCondo = 138039 'PEAC 20071011 reporte de creditos con saldos condonados
    
    gColRecRepCarteraJudCas = 138015 'PEAC 20071022 cartera cred jud y cas
    
    '********************************************************************************
    
    '***By Capi Planeamiento
    gColCredRepFepInforme01 = 108901
    gCapCredRepFepInforme02 = 108902
    gColCredRepFepInforme03 = 108903
    gColCredRepFepInforme3a = 108904
    gColCredRepFepInforme3b = 108905
    gColCredRepFepInforme3c = 108906
    gColCredRepFepInforme3d = 108907
    gColCredRepFepInforme04 = 108908
    gColCredRepFepInforme06 = 108909
    gColCredRepFepEntorno = 108910
    
    '***End By
    
    '**DAOR 20070814*************************************************
    gColCredRepResSalCarxAnaConsolida = 108716 'Resumen de saldo de cartera por analista, de fin de mes
    gColCredRepConCarxAnalistaConsolida = 108717 'Consolidado de cartera por analista, de fin de mes
    '****************************************************************
    gColCredRepSeguroDesgravConsolida = 108718 'DAOR 20071210, Reporte de Seguro de Desgravamen
    gColCredCorresponsaliaPorDebito = 108389 'CORRESPONSALIA A CLIENTES POR DÉBITO 'WIOR 20140530
    gColCredRepCorrespClientes = 108390 'Correspondencia a Clientes
    gColCredRepCartaCobMoro1 = 108391 'Carta de Cobranza para Morosos 1
    gColCredRepCartaCobMoro2 = 108392 'Carta de Cobranza para Morosos 2
    gColCredRepCartaCobMoro3 = 108393 'Carta de Cobranza para Morosos 3
    gColCredRepCartaCobMoro4 = 108394 'Carta de Cobranza para Morosos 4
    gColCredRepCartaCobMoro5 = 108395 'Carta de Cobranza para Morosos 5
    gColCredRepCartaCobMoro6 = 108396 'Carta de Cobranza para Morosos 6 'ARCV 15-02-2007
    gColCredRepCartaInvCredAlt = 108397 'Carta de Invitación Crédito Alterno
    gColCredRepCartaRecup = 108398 'Carta para Recuperaciones
    gColCredRepCartaRecup2 = 108399 'Carta para Recuperaciones 2 AGREGADO POR WIOR 20130910
    gColCredRepCredPersonales = 108400 'Creditos Personales
    gColCredRepCredVig = 108401 'Créditos Vigentes
    gColCredRepCredxInst = 108402 'Créditos por Institución
    gColCredRepMoraxInst = 108403 'Mora por Instituciones
    gColCredRepResSalCartxAna = 108404 'Resumen de Saldos de Cartera x Analista
    gColCredRepResSaldeCartxInst = 108405 'Resumen de Saldos de Cartera x Institución
    gColCredRepLisDesctoPlanilla = 108406 'Listado para Descuento por Planilla
    gColCredRepReportesContabi = 108500 'Reportes para Contabilidad
    gColCredRepPagosconCheque = 108501 'Pagos con Cheque
    gColCredRepPagosdeOtrasAgen = 108502 'Pagos De Otras Agencias
    gColCredRepPagosEnOtrasAgen = 108503 'Pagos En Otras Agencias
    gColCredRepIntEnSusp = 108504 'Intereses en Suspenso
    gColCredRepIntCapit = 108505 'Intereses Capitalizados
    gColCredRepCredFinan = 108506 'CREDITOS Refinanciados
    gColCredRepCredAdjuContabi = 108509 'ALPA***REPORTE DE ADJUDICADOS-CONTABILIDAD
    gColCredRepReportesConsolidados = 108600 'Reportes Consolidados
    gColCredRepConsColocxAgencia = 108601 'Cuadro de Colocaciones x Agencia
    gColCredRepConsColocxAnalista = 108602 'Cuadro de Colocaciones x Analista
    gColCredRepConsMetAlcanzxAna = 108603 'Metas Alcanzadas x Analista
    gColCredRepConsCartAltoRiesgoxAna = 108604 'Cartera de Alto Riesgo x Analista
    gColCredRepConsColocxFteFinan = 108605 'Cuadro de Colocaciones x Fuente Financ
    gColCredRepConsColocyMoraxAna = 108606 'Cuadro de Colocaciones y Mora x Analista
    gColCredRepConsResCartSuper = 108607 'Resumen Cartera - Supervisión
     
    gColCredRepFinMes = 108700                            ' Reportes de Fin de Mes
    gColCredRepFinMesSitCartColocxMoneda = 108701         ' Situacion de Cartera Colocacion x Moneda
    gColCredRepFinMesSitCartColocConsolida = 108702       ' Situacion de Cartera Colocaciones Consolidada
    gColCredRepFinMesResuxProdVenc = 108703               ' Resumen por Producto y Vencimiento
    gColCredRepFinMesResuxProdAge = 108704                ' Resumen por Producto y Agencia
    gColCredRepFinMesListReclasCartera = 108705           ' Listado para Reclasificacion de Cartera
    gColCredRepFinMesIntDeven0_30diaMora = 108706         ' Interes Devengado de 0 - 30 dias de Mora
    gColCredRepFinMesCredVencCobrJudi = 108707            ' Ineteres Devengados Cred Vencidos y Cobranza Judicial
    gColCredRepFinMesResumenGarant = 108708               ' Resumen de Garantias
    gColCredRepFinMesCarteraAltoRiesgo = 108709           ' Cartera Alto Riesgo
    gColCredRepFinMesColocxSectorEcono = 108710           ' Colocacion x Sector Economica Anexo(3)
    gColCredRepFinMesClasifcaIntCred = 108711             ' Clasificacion interes Creditos
    gColCredRepFinMesReporProvCred = 108712               ' Reporte de Provision de Creditos
    gColCredRepFinMesResumenxProdAgenIntDeven = 108713    ' Resumen por Producto y Agencia (Interes Devengado A2.2)
    gColCredRepFinMesResumenxProdAgenIntSuspe = 108714    ' Resumen por Producto y Agencia (Interes Suspenso A2.2)
    gColCredRepFinMesCarteraColoc = 108715                ' Cartera de Colocaciones
    
    gColCredRepPlaneamiento = 108800                      'Reporte para Planeamiento
    gColCredRepPlaneamientoInfo1 = 108801                 'Reporte Info1
    gColCredRepPlaneamientoInfo2 = 108802                 'Reporte Info2
    gColCredRepPlaneamientoInfo3 = 108803                 'Reporte Info3
    gColCredRepPlaneamientoInfo4 = 108804                 'Reporte Info4
    gColCredRepPlaneamientoInfo6 = 108806                 'Reporte info5
    gColCredRepPlaneamientoInfoBCR = 108808               'Informe Colocaciones para BCR
    
    'MAVM 20100510 ***
    gColCredRepResXAnalistaBPPR = 108732                  'Reporte de Resultado de Analista BPPR
    gColCredRepBonificacionXAnalistaBPPR = 108733         'Reporte de Bonificacion de Analista BPPR
    '***
    
End Enum

'******* Reportes de Pignoraticio
Enum ColocPReportes
    gColPRepTran = 128000
    gColPRepTranRegis = 128001
    gColPRepTranDesemb = 128002
    gColPRepTranAnulad = 128003
    gColPRepTranRenov = 128004
    gColPRepTranCancel = 128005
    gColPRepTranRescat = 128006
    gColPRepTranPrendCondDifer = 128007 ' Prendas en condicion diferidas
    gColPRepTranPrendNueva = 128008
    gColPRepTranPagoSobrant = 128009
    
    gColPRepTranDiaOpeExtor = 128015
End Enum

'**DAOR 20071124,Reportes de Calificación ********************************************
Enum ColocCalificacionReportes
    gColCalifCalificacionCartera = 178108
End Enum
'*************************************************************************************

Enum ImpreMuestra
    gImpreMuestraPrevio = 0
    gImpreMuestraArchivo = 1
    gImpreMuestraImpresora = 2
    gImpreMuestraExcel = 3
End Enum



Enum ColocRecTipoCobranza
    gColRecTipCobJudicial = 1
    gColRecTipCobExtraJudi = 2
End Enum

Enum ColocRecDemandado
    gColRecDemandaSi = 1
    gColRecDemandaNo = 2
End Enum

Enum ColocRecViaProcesal
    gColRecViaProcAbreviado = 1
    gColRecViaProcConocimiento = 2
    gColRecViaProcEjecutivo = 3
    gColRecViaProcNoContencioso = 4
    gColRecViaProcNoEspecificado = 5
End Enum

Enum ColocRecExpedienteEstado
    gColRecExpedienteEstEnProceso = 1
    gColRecExpedienteEstSentencia = 2
    gColRecExpedienteEstAbandono = 3
    gColRecExpedienteEstDesistimiento = 4
    gColRecExpedienteEstTransaccion = 5
End Enum


Enum ColocRecTpoComision
    gColRecTpoComiPorcentaje = 1
    gColRecTpoComiMonto = 2
End Enum


Enum ColocRecGastoEstado
    gColRecGastoEstPendiente = 0
    gColRecGastoEstPagado = 1
    gColRecGastoEstEliminado = 2
End Enum


Enum ColocRecExpedTipoInf
    gColRecExpedTipoInfPetit = 1
    gColRecExpedTipoInfHechos = 2
    gColRecExpedTipoInfFundJur = 3
    gColRecExpedTipoInfMedProb = 4
    gColRecExpedTipoInfDatComp = 5
End Enum

Enum ColocRecOperaciones
    gColRecOpePasoARecup = 130100            'Recup : Paso a Recup
    gColRecOpePagJudSD = 130200          'Recup : Pago Jud Sin Demanda
        gColRecOpePagJudSDEfe = 130201        'Recup : Pago Jud SD Efectivo
        gColRecOpePagJudSDChq = 130202        'Recup : Pago Jud SD Cheque
        gColRecOpePagJudSDVou = 130207        'Recup : Pago Jud SD Voucher 'RIRO20140610 ERS017
        
    gColRecOpePagJudCD = 130300          'Recup : Pago Jud Con Demanda
        gColRecOpePagJudCDEfe = 130301        'Recup : Pago Jud CD Efectivo
        gColRecOpePagJudCDChq = 130302        'Recup : Pago Jud CD Cheque
        gColRecOpePagJudCDVou = 130306        'Recup : Pago Jud CD Voucher ' RIRO20140610 ERS017
        
    gColRecOpePagCast = 130400          'Recup : Pago Castigado
        gColRecOpePagCastEfe = 130401        'Recup : Pago Cast Efectivo
        gColRecOpePagCastChq = 130402        'Recup : Pago Cast Cheque
        gColRecOpePagCastVou = 130406        'Recup : Pago Cast Voucher 'RIRO20140610 ERS017
    
    gColRecOpeCastig = 131200          'Recup : Castigar Credito
        gColRecOpeCastigSinDem = 131201        'Recup : Castigar Cred Sin Demanda
        gColRecOpeCastigConDem = 131202        'Recup : Castigar Cred Con Demanda
    
    gColRecOpePagJudSDEnOtCj = 136100          'Recup : Pago Jud Sin Demanda EOCmac
        gColRecOpePagJudSDEnOtCjEfe = 136101        'Recup : Pago Jud SD Efectivo EOCmac

    gColRecOpePagJudCDEnOtCj = 136200          'Recup : Pago Jud Con Demanda EOCmac
        gColRecOpePagJudCDEnOtCjEfe = 136201        'Recup : Pago Jud CD Efectivo EOCmac
        
    gColRecOpePagJudCastEnOtCj = 136300          'Recup : Pago Cast EOCmac
        gColRecOpePagJudCastEnOtCjEfe = 136301        'Recup : Pago Cast EOCmac


    '**** Operaciones EXTORNOS
    gColRecOpeExtTransfRecup = 139100         'Recup : Extorno Transferencia Recup
    gColRecOpeExtPagRecup = 139200            'Recup : Extorno Pago Recup
    
    'FRHU 20150520 ERS022-2015
    gColRecOpePagTransfFocMacm = 103100
        gColRecOpePagTransfFocMacmEfe = 103101
        gColRecOpePagTransfFocMacmChq = 103102
        gColRecOpePagTransfFocMacmVou = 103103
    'FIN FRHU
End Enum


Enum ColRecReportes
    gColRecRepPagoAbogado = 138001
    gColRecRepPagoCredito = 138002
    gColRecRepPagoPorFecha = 138003
    gColRecRepPagoPorAnalista = 138004
    gColRecRepInformeCreditosNuevos = 138011
    gColRecRepInformeCreditosJudicial = 138012
    gColRecRepCreditosCastigadosFechas = 138013
    gColRecRepCreditosenCJxAnalista = 138014
    gColRecRepCreditosenCJxAbogado = 138021
    gColRecRepCreditosVigentes = 138022
    gColRecRepCreditossegunSaldo = 138023
    gColRecRepExpedientesxAbogado = 138024
    gColRecRepResumenMovCobranza = 138025
    gColRecRepActuacionesProcesales = 138034
    gColRecRepActuacionesProcesales2 = 138035
    gColRecRepCreditosCastigados = 138036
End Enum

Enum ColCFModalidad
    gColCFModalidadDEFielCumplimiento = 1
    gColCFModalidadMontoDifPropio = 2
    gColCFModalidadPorAdelanto = 3
    gColCFModalidadOtorgamientoFinanciero = 4
    gColCFModalidadDerechoAduanero = 5
End Enum


Enum ColocCFOperaciones
    gColCFOpeEmision = 140100            'CF : Emision
    gColCFOpeComision = 140200           'CF : Comision
        gColCFOpeComisEfe = 140201       'CF : Comision Efectivo
    gColCFOpeComisMod = 140202           'CF : Comision Efectivo Modificacion CF, WIOR 20120806
    gColCFOpeCancel = 140500             'CF : Cancelacion
    gColCFOpeDevol = 140600           'CF : Devolucion
    
    '**** Operaciones EXTORNOS
    gColCFOpeExtComis = 149100            'CF : Extorno Comision CF
    
End Enum

'******* Reportes de Carta Fianza
Enum ColocCFReportes
    gColCFRepTran = 148000
    gColCFRepCFEmitida = 148001
    gColCFRepCFVigente = 148002
    gColCFRepCFxVencer = 148003
    gColCFRepCFHonrada = 148004
    gColCFRepCFDevuelta = 148005
End Enum

'======= PIGNORATICIO CAJA LIMA

Enum ColocPigTipoTasacion '8013
    gPigTipoTasacNor = 1
    gPigTipoTasacUsoLin = 2
    gPigTipoTasacRetasac = 3
    gPigTipoTasacRetasacVta = 4
End Enum

Enum ColocPigUbica '8004
    gPigUbicaSinCaja = 0
    gPigUbicaBovOfPrin = 1
    gPigUbicaAgLince = 2
    gPigUbicaAgHigue = 3
    gPigUbicaAgBrena = 4
    gPigUbicaAgMagda = 5
    gPigUbicaAgLarco = 6
    gPigUbicaAgSBorja = 7
    gPigUbicaAgMolina = 8
    gPigUbicaRemOfPrin = 9
    gPigUbicaRemLarco = 10
    gPigUbicaTdaOfPrin = 11
    gPigUbicaTdaLarco = 12
    gPigUbicaTdaEsp = 13
    gPigUbicaClubSoc = 14
    gPigUbicaFeriaHog = 15
    gPigUbicaTransito = 16
    gPigUbicaBovValor = 99
End Enum


Enum ColocPigEtapaRemesas '8016

    gPigEtapaRemProyec = 1
    gPigEtapaRemDespac = 2
    gPigEtapaRemRecepc = 3
    
End Enum

Enum ColocPigSituacionPieza '8012
    gPigSituacionDisponible = 1
    gPigSituacionPendFacturar = 2
    gPigSituacionFacturado = 3
    gPigSituacionVenta = 4
    gPigSituacionRescatado = 5
    gPigSituacionFundido = 6
End Enum

Enum ColocPigTipoDocumento '8011
    gPigTipoPoliza = 1
    gPigTipoBoleta = 2
    gPigTipoFactura = 3
End Enum

Enum ColocPigTipoValorMaterial '8008
    gPigTipoValorPrestamo = 1
    gPigTipoValorRemate = 2
    gPigTipoValorVenta = 3
End Enum

Enum ColocPigTipoVenta '8015
    gPigTipoVentaATerceros = 1
    gPigTipoVentaPorResponsabilidad = 2
    gPigTipoVentaAlTitular = 3
End Enum

Enum ColocPigTipoProceso '8009
    gPigTipoPrimeraSubasta = 1
    gPigTipoSegundaSubasta = 2
    gPigTipoTerceraSubasta = 3
    gPigTipoAdjudicacionCliente = 4
    gPigTipoAdjudicacionCaja = 5
    gPigTipoAdjudicacionTasador = 6
    gPigTipoVentas = 7
    gPigTipoFundicion = 8
End Enum

Enum GarProcesos '3501
    gGarProcesosLevantamiento = 901
    gGarProcesosAmpliacion = 902
End Enum

Enum ColocParam '3029
    gColocParamFteIngresoTiempoEval = 1
    gColocParamMiVivDcto = 121
End Enum

Enum ColocRFA
    gColoOpePagNorRFA = 100209
End Enum

Enum ColocITFDesembolso
    gColocITFDesembolso = 990106
End Enum

Public Enum CGTipoIF
    gTpoIFBanco = 1
    gTpoIFFinanciera = 2
    gTpoIFCmac = 3
    gTpoIFCrac = 4
    gTpoIFFuenteFinanciamiento = 5
    gTpoIFCooperativa = 6
    gTpoIFEDPYME = 7
End Enum

Public Enum CGEstadoCtaIF
    gEstadoCtaIFRegistrada = 0
    gEstadoCtaIFActiva = 1    'A
    gEstadoCtaIFCancelada = 2  'B
    gEstadoCtaIFAnulada = 3
    gEstadoCtaIFRestringida = 4
End Enum

Public Enum CGTipoCtaIF
    gTpoCtaIFCtaCte = 1
    gTpoCtaIFCtaAho = 2
    gTpoCtaIFCtaPF = 3
    gTpoCtaIFCtaPFOverNight = 4 'EJVG20120801
    gTpoCtaIFCtaAdeud = 5
    gTpoCtaIFAdeudProyecta = 6
End Enum

Public Enum CGTipoEstCuotaAdeud
    gTpoEstCuotaAdeudPend = 0
    gTpoEstCuotaAdeudCanc = 1
    gTpoEstCuotaAdeudOtros = 2
End Enum

Public Enum CGTipoPagoAdeud
    gTpoPagoAdeudCapital = 1
    gTpoPagoAdeudInteres = 2
    gTpoPagoAdeudGastos = 3
End Enum

Public Enum CGTipoCuotCalIF
    gCGTipoCuotCalIFDesem = 1
    gCGTipoCuotCalIFCuota = 2
    gCGTipoCuotCalIFNoConcesional = 6
End Enum
Public Enum CGEstadosChq
    gCGEstadosChqRecibido = 0
    gCGEstadosChqDepositado = 1
    gCGEstadosChqRechazado = 2
End Enum

Public Enum MovEstado
    gMovEstContabMovContable = 10
    gMovEstContabPendiente = 11
    gMovEstContabRechazado = 12
    gMovEstContabNoContable = 13
    gMovEstContabSustPendRendir = 14
    gMovEstPresupRechazado = 14
    gMovEstPresupPendiente = 15
    gMovEstPresupAceptado = 16
    gMovEstPresupProyectado = 17
    gMovEstLogIngBienAceptado = 20
    gMovEstLogIngBienRechazado = 21
    gMovEstLogSaleBienAlmacen = 22
End Enum

Public Enum MovFlag   ' "gsMovFlag"
    gMovFlagVigente = 0
    gMovFlagEliminado = 1   ' "X"
    gMovFlagExtornado = 2   '  "E"
    gMovFlagDeExtorno = 3   ' "N"
    gMovFlagDeLeido = 4     '  "L"
    gMovFlagModificado = 5  ' "M"
End Enum

Public Enum MovParalelo   '= "CMovParalelo"
    gMovParaleloTransferido = 1   ' "T"
End Enum

Public Enum CHTipoProc   '    Tipo de Proceso de Caja Chica
    gCHTipoProcHabilitacion = 1     'Habilitacion
    gCHTipoProcRendicion = 2   'Rendición en Contabilidad
    gCHTipoProcArqueo = 3  'Arqueo
    gCHTipoProcDesembolso = 4  'Desembolso de Caja Chica
    gCHTipoProcEgresoDirecto = 5    'Egreso Directo de Caja Chica
    gCHTipoProcArendir = 6     'A rendir Cuenta
    gCHTipoProcNuevaPorRendicion = 7 'Nueva por rendicion
    gCHTipoProcCancelada = 8 'Cancelada
End Enum

Public Enum ArendirTipo  '     Tipo de Arendir Cuenta
    gArendirTipoCajaGeneral = 1   'A rendir Cuenta de Caja General
    gArendirTipoViaticos = 2     'A rendir Cuenta de Viáticos
    gArendirTipoCajaChica = 3    'A rendir Cuenta de Caja Chica
    gArendirTipoAgencias = 4     'A rendir Cuenta de Agencias
End Enum

Public Enum ViaticosCateg         'Viaticos Categoría del Trabajador CMACT
    gViaticosCategFuncionario = 1  ' A     'Funcionario
    gViaticosCategEmpleado = 2     ' B    'Empleado
End Enum

Public Enum ViaticosAfectoA
    gViaticoAfectoADias = 1     '  D   Nro de Dias
    gViaticoAfectoATransporte = 2   ' T   Transporte
    gViaticoAfectoANinguno = 3      'N   Ninguno
End Enum

Public Enum ViaticosTransporte         'Viaticos Transporte
    gViaticosTransporteAereo = 1    'Aereo
    gViaticosTransporteMaritimo = 2     'Maritimo
    gViaticosTransporteTerrestre = 3    'Terrestre
End Enum

Public Enum ViaticosDestino            'Viaticos Destino
    gViaticosDestinoLaLibertad = 1        'La Libertad
    gViaticosDestinoLimaIqAreqCuz = 2     'Lima, Iquitos, Arequipa o Cuzco
    gViaticosDestinoOtros = 3             'Otras Ciudades
End Enum

Public Enum OpeTpoVisible
   gOpeTpoVisibleSi = 1   '1 SI
   gOpeTpoVisibleNo = 0   '0 NO
End Enum

Public Enum OpeCtaCaracter
   gOpeCtaCaracOpcional = 2      'Opcional
   gOpeCtaCaracObligatorio = 1   'Obligatorio
End Enum

Public Enum FiltroObj
    CtaOBjFiltroObj = 0
    CtaOBjFiltroIF = 1
End Enum
Public Enum TpoObjetos
    objPersona = 0
    ObjEntidadesFinancieras = 1
    ObjAccionesSectorPrivado = 2
    ObjCMACAgencias = 11
    ObjCMACArea = 12
    ObjCMACAgenciaArea = 13
    ObjBienesServicios = 18
    ObjGastosOrgInstalacion = 19
    ObjConceptosARendir = 20
    ObjArendir = 50
    ObjCtasAhorrosCMACS = 51                   'CUENTAS DE AHORROS DE CMACT
    ObjProductosCMACT = 60
    ObjDescomEfectivo = 90
    ObjServRecTerceros = 4511
    ObjGastosDirectorio = 4415                 'GASTOS DE DIRECTORIO
    ObjOtrasCargasDivGestion = 4719            'OTRAS CARGAS DIVERSAS DE GESTION
End Enum
Public Enum TpoMargenes
    MargenSuperior = 0
    MargenInferior = 1
    MargenDerecho = 2
    MargenIzquierdo = 3
End Enum
Public Enum OpeDocEstado
    OpeDocEstObligatorioDebeExistir = 11
    OpeDocEstObligatorioNoDebeExistir = 12
    OpeDocEstOpcionalDebeExistir = 21
    OpeDocEstOpcionalNoDebeExistir = 22
End Enum
Public Enum OpeDocMetodo
    OpeDocMetNroMovimiento = 1
    OpeDocMetAutogenerado = 2
    OpeDocMetDigitado = 3
End Enum
Public Enum ARendirFases
    ArendirSolicitud = 0
    ArendirRechazo
    ArendirAtencion
    ArendirSustentacion
    ArendirRendicion
    ArendirExtornoAtencion
    ArendirExtornoRendicion
End Enum

Public Enum CGEstadosOp
    PorEmitir = 0
    Emitidas = 1
    Anuladas = 2
    Rechazadas = 3
    Extraviadas = 4
End Enum

Public Enum CGEstadoConfCheque
    ChqCGSinConfirmacion = 0
    ChqCGNoConfirmado = 1
    ChqCGConfirmado = 2
End Enum
Public Enum CGTipoTransporte
    CGTipoTransporteBlindado = 0
    CGTipoTransportePropio = 1
    CGTipoTransporteAlquilado = 2
End Enum

Public Enum CGCtaIFConceptos
    CGCtaIFConCapital = 1
    CGCtaIFConInteres = 2
    CGCtaIFConIntAdel = 3
    CGCtaIFConMora = 4
    CGCtaIFConGastos = 5
    CGCtaIFConGastosServ = 6
    CGCtaIFConComisiones = 7
End Enum
Public Enum CGBloqueos
   gBloqueoBalance = 1
   gBloqueoSaldos = 2
End Enum


Public Enum CGAnxTipoCalculo
   gAnxTCalcTasa = 1
   gAnxTCalcVencimiento = 2
   gAnxTCalcEncajeBCR = 3
   gAnxTCalcPlazaCheque = 4
   gAnxTCalcSeries = 5
   gAnxTCalcFecha = 6
   gAnxTCalcFormula = 7
   gAnxTCalcFormulasAcumula = 8
   gAnxTCalcTotales = 9
End Enum

Public Enum CGAdeudCalTpoCuota
   gAdeudTpoCuotaFija = 1
   gAdeudTpoCuotaVariable = 2
End Enum

Public Enum ContAnx10TipoOficina
    gAnx10TipoOfPrincipal = 1
    gAnx10TipoOfSucursal = 2
    gAnx10TipoOfAgencia = 3
    gAnx10TipoOfEspecial = 4
    gAnx10TipoOfLocalCompar = 5
End Enum

Public Enum ARendirDestinoGasto
    gARendirDestinoGastoJudiciales = 1
    gARendirDestinoGastoOtros = 2
    gARendirDestinoGastoViaticos = 3
    gARendirDestinoGastoNoDefinido = 4
End Enum

Public Enum ConstanteCabecera
    gProducto = 1001
    gTpoProducto = 3033 'MAVM 20100604 BAS II
    gTpoCredito = 3034 'ALPA 20100605 BAS II
    gTpoInstFinanc = 3035 'MAVM 20100604 BAS II
    gColRelacPersInfoGas = 3036 'BRGO 20111104
    
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
    
    gTipoContratoGarantR = 1051
    gTipoPolizaAlternativaInmobil = 1052 'EJVG20150612
    gTipoPolizaAlternativaMobil = 1056 'EJVG20150612
    gTipoPinPad = 9050
    
    gTipoPuertoSerial = 1037
    gPersEstLavDinero = 1038
    'By Capi 30042008
    gPersCondControl = 9072
    '
    gPersTpoSangre = 1039
    gGarantTpoInmueb = 1040
    gGastosProd = 1041
    gChequeTipo = 1042
    gAreasAgrupacion = 1043  'Agrupacion de Areas
    gTipoExoneracionITF = 1044  'Agrupacion de Areas
    gTipoBienContrato = 1063 'CTI3 ERS0012020
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
    gServConcepto = 2025
    gCapTasaEspEstados = 2026
    gCaptacMotBloqueoParcial = 2029
    gCaptacSubProdAhorros = 2030
    gCapNivRetCancEstado = 2031
    gCaptacSubProdPlazoFijo = 2032
    gCaptacSubProdCTS = 2033
    gCapEstAutorizaOpe = 9051  'ppoa
    
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
    gColocTipoPagoEspecial = 3021 'JUEZ 20150415
    gColocMetasCred = 3022
    gColocTipoMetas = 3023
    gColocMotivRechazo = 3024
    gColocGarantRec = 3025
    gColocGarantRecEstado = 3026
    'gColCFModalidad = 3027
    gColocMiVivTramo = 3028
    'MAVM 20120606 ***
    gColocMiVivPorcentCI = 3060
    gColocMiVivValVentMin = 3061
    gColocMiVivValVentMax = 3062
    gColocMiVivPlazoMin = 3065
    gColocMiVivPlazoInt = 3069 'MIVIVIENDA: Monto a Financiar Minimo Tramo 2    125000.00
    gColocMiVivPlazoMax = 3066
    gColocMiVivBonoBuenPagador = 3067
    gColocMiVivBonoBuenPagador2 = 3068 '3068    BONO FAMILIAR HABITACIONAL 2    5000.00
    
    

    '***
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
    gColocRecTipoProceso = 3307
    gColocRecTipoProcesal = 3308
    gColocRecExpedienteEstado = 3309
    gColocRecTipoActProcesal = 3310 'DAOR 20070124
    gColocRecTipoMedCautelar = 3311 'DAOR 20070417
    'CF
    gColCFModalidad = 3402
    gColocCFOperaciones = 3403
    gColocCFReportes = 3403
    gColocCFMotivHonrar = 3404
    
    'Garantias
    gGarProcesos = 3501
    
    'CTI06    20201208
    gCVTipoPagoBase = 5020
    
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
    
    'By capi 28102008
    gColocMotivoRef = 3032
    '
    gSegTarjetaTipoSiniestro = 1200 'FRHU 20140515 ERS068-2014
    gSegTarjetaTipoDocumento = 1201 'FRHU 20140515 ERS068-2014
    gSegTarjetaEstadoActivacion = 1203 'FRHU 20140515 ERS068-2014
    'EJVG20150123 ***
    gGarantiaTpoClasif = 7090
    gGarantiaTpoSinTramite = 7091
    gGarantiaTpoValorizacion = 7092
    gGarantiaClaseTasacInmobiliaria = 7093
    gGarantiaCategoriaTasacInmobiliaria = 7094
    gGarantiaTramiteLegalTipo = 7104 'FRHU 20150428 Cambio x 7098
    gPolizaTipo = 7106 'EJVG20150608
    gCompraDeudaDestino = 7108 'EJVG 20160702 ERS002-2016
    gGarantiaOficinaRegistral = 7099
    gGarantiaZonaRegistral = 7100
    gGarantiaTramiteLegalEstado = 7101
    gGarantiaTramiteLegalCobertura = 7102
    'END EJVG *******
End Enum

'MAVM 20110523 ***
Public Enum ColocCredLlaveProceso
    gColCredNivApr = 9983
End Enum
'***

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
    
    gColPYMEConv = 204 'MAVM 20091028 Nuevo Producto
    
    gColConsuDctoPlan = 301
    gColConsuPlazoFijo = 302
    gColConsCTS = 303
    gColConsuUsosDiv = 304
    gColConsuPrendario = 305
    gColConsPFTpoProducto = 703 'MAVM 20111223
    gColPigTpoProducto = 705 'MAVM 20100604 BAS II ***
    gColConsuPrestAdm = 320
    gColCFComercial = 121
    gColCFPYME = 221
    gColHipoCaja = 401
    gColHipoMiVivienda = 423
    
    '*** PEAC 20080822 *****************
    gColAperturaSolicitudCred = 100914
    gColRegistraAprobacion = 100915
    gColRegistraContratoPig = 100916
    '*** FIN PEAC **********************
    'BASILEA II****************************************-
    gColCFTpoProducto = 514 'MAVM 20100604 BAS II ***
    
    'Tipo credito
    'ALPA 20100604************************
    gColCredCorpo = 150
    gColCredCorpoCF = 181 'MAVM 20100609 BAS II
    gColCredGrEmp = 250
    gColCredMeEmp = 350
    gColCredPeEmp = 450
    gColCredMicro = 550
    gColCredConRE = 650
    gColCredConNR = 750
    gColCredHipot = 850
    
    gColProConsumo = 700
    gColProConsumoPerDesPla = 704
    gColProConsumoPrendario = 705
    gColProHipoteMiVivienda = 801
    gColProHIpoConveTrayDir = 805   'ALPA 20110128
    '***************************************************
    
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
    gPersonaJurCFLFinancieras = 9 '***Agregado por ELRO el 20111105, según Acta 307-2011/TI-D
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
    gPersIdREPEV = 16 '"REPEV(PARA TASADORES)"'WIOR 20140703


    '20200825-LARI DOCUMENTOS PARA ENTRANGEROS SEGÚN OFICIO 13323-2020-SBS****************************
    gPersIdCIEMRE = 17 '17  CIEMRE         (CARNET DE INDENTIDAD EMITIDO POR EL MINISTERIO DE RELACIONES EXTERIORES)
    gPersIdCPTP = 18   '18  CPTP           (CARNET DE PERMISO TEMPORAL DE PERMANENCIA)
    gPersIdCI = 19     '19  CI           (CÉDULA DE IDENTIDAD O DOCUMENTOS ANÁLOGOS)
    gPersIdCREF = 20   '20  CARNET DE REFUGIADO
    gPersIdCEPR = 21   '21  CEPR           (DOCUMENTO EXPEDIDO POR LA CEPR DEL MINISTERIO DE RELACIONES EXTERIORES)17  CIEMRE         (CARNET DE INDENTIDAD EMITIDO POR EL MINISTERIO DE RELACIONES EXTERIORES)
    '*************************************************************************************************
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
    gPersTipoConvenioRecaudo = 10 'Modificado por RIRO el 20130314
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
    'ALPA**20080812*
    gPersGarantEstadoRecuperado = 7
    gPersGarantEstadoAdjudicado = 8
    gPersGarantEstadoRematado = 9
    gPersGarantEstadoVendido = 10
    '***************
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
    TCPondVenta = 7
    TCPondREU = 8
    TCSBSDia = 9 'ALPA20140225****************************
    TCCompraTraider = 10 'ALPA20140225****************************
    TCVentaTraider = 11 'ALPA20140225****************************
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
    gGruposIngRecaudo = 9 ' Agregado por RIRO el 20130318
    gGruposIngEgreNoEfectivo = 10 'FRHU 20141203 ERS048-2014
    gGruposTransferencia_CCE = 11 'ALPA 20161204 CCE
    
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
    TpoDocCtaCargoApert = 150 'JUEZ 20131212
    TpoDocCtaCargoAumCap = 151 'JUEZ 20131212
    TpoDocCtaCargoPagoCred = 152 ' JUEZ 20131212
    'ALPA 20090320*****************************
    TpoDocRetenciones = 0
    '******************************************
End Enum

Public Enum TipoCalculoInteres
    TpoCalcIntSimple = 0
    TpoCalcIntCompuesto = 1
    TpoCalcIntAdelantado = 2
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
    gConstSistMostrarRegistroEfectivo = 85 'FRHU 20160225
    gConstSistOpeTpoMostrarRegistroEfectivo = 86 'FRHU 20160225
    gConstSistIntervinientesDelCredito = 87 'FRHU 20160615 ERS002-2016
    gConstSistCondicionesDelCreditoAvalidar = 88 'FRHU 20160615 ERS002-2016
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
    
    gConstSistPasswordHojaExcel = 180 'FRHU 20141218 OBSERVACION
    gConstSistRHCargoCodPromotores = 327
    gConstSistRHCargoCodPosicionCli = 328
    gConstSistMinimoRatioCobertura = 355 'FRHU 20160811 Anexo002 ERS002-2016
    gConstSistDiaPagoCuotaEcotaxi = 419 'EJVG20120521
    gConstSistMontoPenalidadEcoTaxixAtrasoEntregaxDia = 420 'EJVG20120621
    gConstSistMontoPenalidadEcoTaxixVehiculoDistintoAcordado = 421 'EJVG20120621
    gConstSistRHCargoCodCTSNoActiva = 424 '***Agregado por ELRO el 20121015, según OYP-RFC101-2012
    gConstSistRHCargoCodAgenciasCredEval = 425 'Juez 20120907
    gConstSistRHCargoCodAnalistasCredEval = 426 'Juez 20120907
    gConstSistRHCargoCodJefeAgenciaNegocio = 561 'JOEP 20170123
    gConstSistRHCargoCodRiesgos = 562 'JOEP 20170123
    gConstSistRHCargoCodCoordCredEval = 427 'Juez 20120907
    gConstSistRHCargoCodJefCredEval = 428 'Juez 20120907
    gConstSistRHCargoCodModAgeManProPas = 430 '***Agregado por ELRO el 20121015, según OYP-RFC123-2012
    gConstSistAgenciaHabNivelesApr = 433 'JUEZ 20130411
    gConstSistDiaPagoCalendEcotaxi = 434 'EJVG20120521
    gConstSistDiaPagoCuotaEcotaxiCoberturaOG = 435 'EJVG20120521
    gConstSistDiaPagoCalendEcotaxiCoberturaOG = 436 'EJVG2012052
    gConstSistEstadoCampActualizaDatos = 445 'JUEZ 20131022
    gConstSistRHCargoCodMantCampDatos = 446 'JUEZ 20131022
    gConstSistFechaBPP = 438 'WIOR 20131010
    gConstSistBitPolizaManual = 509 'EJVG20150702
    
    gConstSistCuotaSistemaFinanciero = 528 'ALPA20150623
    gConstSistCuotaTarjetaSistFinanc = 529 'ALPA20150623
    
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

'AMDO TI-ERS086-2013 20130726
  Public Enum PermisoCargos
    gPerPosCliente = 1
    gMantProdPasivos = 2
    gEdicionPersonas = 3 'WIOR 20130826
    gPlanDePagos = 4 'FRHU 20140722 ERS105-2014
    gResumenIngEg = 5 'RIRO20140930 ERS115
    gMantParamLavado = 6 'JUEZ 20141124
    gMantParamCreditos = 7 'JUEZ 20141124
    gMantParamAhorros = 8 'JUEZ 20141124
    gMantParamLibreImpConsMov = 9 'JUEZ 20151229
    gMantParamComisiones = 10 'JUEZ 20151229
  End Enum
'END AMDO

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
    gClaveNOValida = -1
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



Public Enum RHBeneficiarioRela
    gPersRelaBenefDesaCtivado = 0
    gPersRelaBenefACtivo = 1
End Enum

Public Enum RHAsistenciaMedicaRela
    gPersRelaAMPDesactivado = 0
    gPersRelaAMPActivo = 1
End Enum

Public Enum RHProcesoSeleccionTipo
    gRHProcSelTpoInterno = 0
    gRHProcSelTpoExterno = 1
End Enum

Public Enum RHProcesoSeleccionEstado
    gRHProcSelEstIniCiado = 0
    gRHProcSelEstFinalizado = 1
End Enum

Public Enum RHProcesoSeleccionResultado
    gRHProcSelResIngreso = 0
    gRHProcSelResSuplente = 1
    gRHProcSelResNoIngreso = 2
End Enum

Public Enum RHEstado
    gRHEstInactivo = 0
    gRHEstActivo = 1
    gRHEstVacacFisicas = 2
    gRHEstVacacSuspend = 3
    gRHEstLicSGoceHaber = 4
    gRHEstLicCGoceHaber = 5
    gRHEstSubs = 6
    gRHEstSuspend = 7
    gRHEstDespedido = 8
    gRHEstRetirado = 9
End Enum

Public Enum RHCondicion
    gRHCondEmpContratado = 0
    gRHCondEmpEstable = 1
End Enum

Public Enum RHCategoria
    gRHCategFunCionario = 1
    gRHCategEmpleado = 2
End Enum

Public Enum RHCargoEstado
    gRHCargoEstACtivo = 1
    gRHCargoEstEliminado = 2
End Enum

Public Enum RHConceptoIngresos
    gRHConcIngBoncarFam = 101
    gRHConcIngBonConsol = 102
    gRHConcIngBonVacacional = 103
    gRHConcIngBonCargo = 104
    gRHConcIngBonResponsabilidad = 105
    gRHConcIngCTSBonVac = 106
    gRHConcIngCTSBonGratif = 107
    gRHConcIngCTSTercio = 108
    gRHConcIngAFP1023 = 109
    gRHConcIngAFP0300 = 110
    gRHConcIngIngDiaSub = 111
    gRHConcIngNetoPagar = 112
    gRHConcIngONP = 113
    gRHConcIngPrueba1 = 114
    gRHConcIngReintegro = 115
    gRHConcIngRemNoAfeCta = 116
    gRHConcIngRemVacacional = 117
    gRHConcIngRemMensual = 118
    gRHConcIngSueBasico = 119
    gRHConcIngTerBonCargo = 120
    gRHConcIngTerBonCarFam = 121
    gRHConcIngTerBonConsol = 122
    gRHConcIngTerCio = 123
    gRHConcIngTerAFP1023 = 124
    gRHConcIngTerAFP300 = 125
    gRHConcIngTerSueBasico = 126
    gRHConcIngTerBonResp = 127
    gRHConcIngTotBonProduct = 128
    gRHConcIngTotCTS = 129
    gRHConcIngTotIngresos = 130
    gRHConcIngTotalUtilidades = 131
    gRHConcIngTotalGratifiCaCiones = 132
    gRHConcIngBonProduct = 133
    gRHConcIngIngresosLocacion = 134
    gRHConcIngAguinaldoNavideño = 135
    gRHConcIngAguinaldoJuguetes = 136
    gRHConcIngVacSueldoBasico = 137
    gRHConcIngVacBonCargo = 138
    gRHConcIngVacAFP300 = 139
    gRHConcIngVacBonResponsabilidad = 140
    gRHConcIngVacBonFam = 141
    gRHConcIngVacAFP1023 = 142
    gRHConcIngVacCTSTotRem = 143
    gRHConcIngVacBonConsol = 144
    gRHConcIngTerBonProduCtividad = 145
    gRHConcIngVacSueBasico = 146
End Enum

Public Enum RHConceptoDescuentos
    gRHConcDcto5taCat = 201
    gRHConcDctoAFPComVar = 202
    gRHConcDctoAFPCuotaFija = 203
    gRHConcDctoAFPPrima = 204
    gRHConcDctoAMP = 205
    gRHConcDctoBcoLima = 206
    gRHConcDctoEsSaludVida = 207
    gRHConcDctoJudFijo = 208
    gRHConcDctoJudPorc = 209
    gRHConcDctoHipotec = 210
    gRHConcDctoTardanzas = 211
    gRHConcDctoONP1300 = 212
    gRHConcDctoPrueba = 213
    gRHConcDctoPrestAdmin = 214
    gRHConcDctoTotalDctos = 215
End Enum

Public Enum RHConceptoAportaciones
    gRHConcAportExtSolid = 301
    gRHConcAportSegSoc = 302
    gRHConcAportTotalAport = 303
End Enum

Public Enum RHConceptoVariablesUsuario
    gRHConcVarUsuCTSMesTrab = 401
    gRHConcVarUsuGratMesTrab = 402
    gRHConcVarUsuSubNumDias = 403
    gRHConcVarUsuPorcMesTrab = 404
End Enum

Public Enum RHConceptoVariablesGlobales
    gRHConcVarGlobTotDiasLab = 501
    gRHConcVarGlobPorcPagUTIL = 502
    gRHConcVarGlobSumTotIng = 503
    gRHConcVarGlobValUtil = 504
End Enum

Public Enum RHConceptoFuncionesConstantes
    gRHConcFuncConstAFPValComVar = 601
    gRHConcFuncConstAFPValorPrima = 602
    gRHConcFuncConstDctoJudiFijo = 603
    gRHConcFuncConstDctoJudiPorc = 604
    gRHConcFuncConst5taCat = 605
    gRHConcFuncConstAFP1023 = 606
    gRHConcFuncConstAFP0300 = 607
    gRHConcFuncConstCTSMesTrab = 608
    gRHConcFuncConstMinSal = 609
    gRHConcFuncConstMinTard = 610
    gRHConcFuncConstDiasnoTrab = 611
    gRHConcFuncConstNumMesTrab = 612
    gRHConcFuncConstTERNumMesTrab = 613
    gRHConcFuncConstPagoExeso = 614
    gRHConcFuncConstReintegro = 615
    gRHConcFuncConstSdoBas = 616
    gRHConcFuncConstSdoBasNivel = 617
    gRHConcFuncConstSdoContrato = 618
    gRHConcFuncConstUltBonVacac = 619
    gRHConcFuncConstUltGratif = 620
    gRHConcFuncConstUltTercio = 621
End Enum

Public Enum RHConceptoVariablesLocales
    gRHConcVarTotIngAcum = 701
    gRHConcVarAFPCuotaFija = 702
    gRHConcVarIngAntesAFP = 703
    gRHConcVarMontoAMP = 704
    gRHConcVarNumDiasLab = 705
    gRHConcVarNumDiasMes = 706
    gRHConcVarNumMinTrab = 707
    gRHConcVarPrueba = 708
    gRHConcVarPorcAsistMedFam = 709
    gRHConcVarPorcONP = 710
    gRHConcVarPorcFonavi = 711
    gRHConcVarPorc5taCat = 712
    gRHConcVarPorcSegSoc = 713
    gRHConcVarPorcExtSolid = 714
    gRHConcVarPorcIncONP = 715
    gRHConcVarPorcRPS = 716
    gRHConcVarPorcTercio = 717
    gRHConcVarPromDia = 718
    gRHConcVarSueldoNeto = 719
    gRHConcVarSumMonInt = 720
    gRHConcVarAFPTopePrima = 721
    gRHConcVarAPFTope = 722
    gRHConcVarTotalRemun = 723
    gRHConcVarUITValor = 724
    gRHConcVarUIT54 = 725
    gRHConcVarUIT07 = 726
    gRHConcVarDiaProm = 727
End Enum

Public Enum RHConceptoTipo
    gRHTpoConcConst = 1
    gRHTpoConcValCal = 2
    gRHTpoConcValPreDef = 3
    gRHTpoConcFunInt = 4
End Enum

Public Enum RHConceptoEstado
    gRHConcEstAct = 1
    gRHConcEstDesh = 0
End Enum

Public Enum RHImpuesto5taCat
    gRHImp5taCatCDcto = 1
    gRHImp5taCatSDcto = 0
End Enum

Public Enum RHAfectoMesTrabajado
    gRHAfeCtoMesTrabajado = 1
    gRHNoAfeCtoMesTrabajado = 0
End Enum

Public Enum RHProcesoSeleccionModal
    gRHProcSelModInterno = 0
    gRHProcSelModNuevo = 1
End Enum

'**************************************
'           CONSTANTES DE CREDITOS
'**************************************

Public Enum CredTipoFiltro
    'Filtro de Constantes
    gCredFiltroProd = 1 'Para Filtrar Todos los Productos de Creditos
    gCredFiltroProdCab = 2 'Para Filtrar Todos los Cabeceras de Productos de Creditos
    gCredFiltroCalendApl = 3 'Para Filtrar Aplicacion de Calendario A Cuota o Desembolso
    gCredFiltroInstRepres = 4 'Para Filtrar Representantes de Institucion
    gCredFiltroGastosOperad = 5 'Para filtrar operaciones de Comisiones
End Enum

Public Enum CredOperacion
    
    'Garantias
    gsGarantLevanta = 160101
    
    
    'Aprobacion de Credito
    gCredAprobacion = 101501
    
    'Desembolso en Efectivo
    gCredDesembEfec = 100101
    'Desembolso Abono a Cuenta Nueva
    gCredDesembCtaNueva = 100102
    'Desembolso Abono a Cuenta Existente
    gCredDesembCtaExist = 100103
    'Desembolso Abono a Cuenta Existente de Otra Agencia
    gCredDesembCtaExistDOA = 100104
    'Desembolso Abono a Cuenta Nueva de Otra Agencia
    gCredDesembCtaNuevaDOA = 100105
    
    'Desembolso Retiro Gastos De Agencia Local
    gCredDesembCtaRetiroGastos = 100106
    'Desembolso Retiro Gastos De Agencia Remota
    gCredDesembCtaRetiroGastosDOA = 100107
    'Desembolso Retiro Cancelacion Credito Agencia Local
    gCredDesembCtaRetiroCancelacion = 100108
    'Desembolso Retiro Cancelacion Credito Agencia Remota
    gCredDesembCtaRetiroCancelacionDOA = 100109
    
    'Desembolso con Cheque
    gCredDesembCheque = 100110
    
    'Desembolso con Leasing
    gCredDesembLeasing = 100111
    
    'Desembolso con Banco de la Nacion
    gCredDesembBN = "100113"
    
    'ALPA 20111209*********************
    gCredPagLeasingCI = "100211"
    gCredPagLeasingCU = "100212"
    gCredPagLeasingPC = "100216"  '20140124*******************
    '**********************************
    gCredPagoCuotasEcotaxi = "100213" 'EJVG20120518
    gCredPagoCuotasEcotaxiCoberturaOG = "100214" 'EJVG20120518
    'gCredPagoCuotasAuto = "100220" 'WIOR 20140620
    gCredPagoCuotasDebitoAuto = "100220" 'JUEZ 20150130
    gCredPagoCuotasLiqCredSegDes = "100225" 'WIOR 20160413
   'Desembolso con InfoGas
    gCredDesembInfoGas = 100112
    
    'Pagos Lote
    gCredPagLote = 102100
    gCredPagNorLote = 102103 'RIRO20141014 ERS122
    
    'Registro de Dacion
    gCredRegisDacion = 102200
    
    'Pagos Normales
    gCredPagNorNor = 100200
    gCredPagNorNorEfec = 100201
    gCredPagNorNorCC = 100202
    gCredPagNorNorCh = 100203
    gCredPagNorNorDctoPla = 100204
    'Pagos de Honramiento
    gCredPagHonramiento = "100218" 'WIOR 20131221
    
    'WIOR 20150421 ***
    'Reversion de Interes Diferido
    gCredRevIntDifVig = 100930
    gCredRevIntDifVenc = 100931
    gCredRevIntDifRef = 100932
    gCredRevIntDifJud = 100933
    'Transferencia de Interes Diferido A Refinanciado
    gCredTransRefIntDifVig = 100936
    gCredTransRefIntDifVenc = 100937
    'Transferencia de Interes Diferido A Recuperaciones
    gCredTransRecupIntDifVig = 100938
    gCredTransRecupIntDifVenc = 100939
    'WIOR FIN ********
    
    
    gCredIntDifeAmpliado = 100940 'VAPA20170817
    'Llamada Reccepcion en Otra Cmact
    gCredPagNorNorEOCEfec = 100205
    gCredPagNorNorEOCCh = 100206
    'Dacion en Pago
    gCredPagNorNorDacion = 100207
    gCredPagNorNorVouch = 100219 'RIRO20140530 ERS017
    gCredPagAnticipAdelantoCuota = 100221 'JUEZ 20150415
    
    'Pagos RFA
    gCredPagNorRfaEfec = 100209
    
    gCredPagNorMor = 100300
    gCredPagNorMorEfec = 100301
    gCredPagNorMorCC = 100302
    gCredPagNorMorCh = 100303
    gCredPagNorMorDctoPla = 100304
    'Llamada Reccepcion en Otra Cmact
    gCredPagNorMorEOCEfec = 100305
    gCredPagNorMorEOCCh = 100306
    'Dacion en Pago
    gCredPagNorMorDacion = 100307
    gCredPagNorMorVouch = 100308 'RIRO20140530 ERS017
    
    gCredPagNorVen = 100400
    gCredPagNorVenEfec = 100401
    gCredPagNorVenCC = 100402
    gCredPagNorVenCh = 100403
    gCredPagNorVenDctoPla = 100404
    'Llamada Reccepcion en Otra Cmact
    gCredPagNorVenEOCEfec = 100405
    gCredPagNorVenEOCCh = 100406
    'Dacion en Pago
    gCredPagNorVenDacion = 100407
    gCredPagNorVenVouch = 100409 'RIRO20140530 ERS017
    
    gCredPagRefNor = 100500
    gCredPagRefNorEfec = 100501
    gCredPagRefNorCC = 100502
    gCredPagRefNorCh = 100503
    gCredPagRefNorDctoPla = 100504
    'Llamada Reccepcion en Otra Cmact
    gCredPagRefNorEOCEfec = 100505
    gCredPagRefNorEOCCh = 100506
    'Dacion en Pago
    gCredPagRefNorDacion = 100507
    gCredPagRefNorVouch = 100508 'RIRO20140530 ERS017
    
    gCredPagRefMor = 100600
    gCredPagRefMorEfec = 100601
    gCredPagRefMorCC = 100602
    gCredPagRefMorCh = 100603
    gCredPagRefMorDctoPla = 100604
    'Llamada Reccepcion en Otra Cmact
    gCredPagRefMorEOCEfec = 100605
    gCredPagRefMorEOCCh = 100606
    'Dacion en Pago
    gCredPagRefMorDacion = 100607
    gCredPagRefMorVouch = 100608 'RIRO20140530 ERS017
    
    gCredPagRefVen = 100700
    gCredPagRefVenEfec = 100701
    gCredPagRefVenCC = 100702
    gCredPagRefVenCh = 100703
    gCredPagRefVenDctoPla = 100704
    'Llamada Reccepcion en Otra Cmact
    gCredPagRefVenEOCEfec = 100705
    gCredPagRefVenEOCCh = 100706
    'Dacion en Pago
    gCredPagRefVenDacion = 100707
    gCredPagRefVenVouch = 100708 ' RIRO20140530 ERS017
    
    'Pago en Suspenso
    gsPagIntSus = 100800
    gsPagIntSusEfec = 100801
    
    'Condonación con el Módulo de Perdón de Mora
    gsPerdonMoraCred = "100904" 'NAGL Según ERS036-2018
    
    'JUEZ 20131022 *********************
    gCredReprogDesastreNat = "100920"
    gCredReprogRoyaAmarilla = "100921"
    'END JUEZ **************************
    
    gCredReprogZonaInundada = "100925" 'PASI 20170405
    
    '*************************************
    '******   Para Refinanciacion ********
    '*************************************
    gCredRefinanciacion = 101200
    gCredRefNormSoles = 101201
    gCredRefMorSoles = 101202
    gCredRefVencSoles = 101203
    gCredRefIngSoles = 101261
    
    gCredRefRefNormSoles = 101204
    gCredRefRefMorSoles = 101205
    gCredRefRefVencSoles = 101206
    gCredRefRefIngSoles = 101207
    
    
    gCredRefNormDolares = 101247
    gCredRefMorDolares = 101248
    gCredRefVencDolares = 101249
    gCredRefIngDolares = 101262
    
    gCredRefRefNormDolares = 101250
    gCredRefRefMorDolares = 101251
    gCredRefRefVencDolares = 101252
    gCredRefRefIngDolares = 101253
    
    gCredRefDMNormSoles = 101208
    gCredRefDMMorSoles = 101209
    gCredRefDMVencSoles = 101210
    gCredRefIngDMSoles = 101263
    
    gCredRefRefDMNormSoles = 101211
    gCredRefRefDMMorSoles = 101212
    gCredRefRefDMVencSoles = 101213
    gCredRefRefIngDMSoles = 101214
    
    gCredRefDMNormDolares = 101215
    gCredRefDMMorDolares = 101216
    gCredRefDMVencDolares = 101217
    gCredRefIngDMDolares = 101264
    
    gCredRefRefDMNormDolares = 101218
    gCredRefRefDMMorDolares = 101219
    gCredRefRefDMVencDolares = 101220
    gCredRefRefIngDMDolares = 101221
    
    gCredRefCapIntNormSoles = 101222
    gCredRefCapIntMorSoles = 101223
    gCredRefCapIntVencSoles = 101224
    gCredRefCapIntIngSoles = 101265
    
    gCredRefRefCapIntNormSoles = 101225
    gCredRefRefCapIntMorSoles = 101226
    gCredRefRefCapIntVencSoles = 101227
    gCredRefRefCapIntIngSoles = 101228
    
    gCredRefCapIntDMNormSoles = 101229
    gCredRefCapIntDMMorSoles = 101230
    gCredRefCapIntDMVencSoles = 101231
    gCredRefCapIntIngDMSoles = 101266
    
    gCredRefRefCapIntDMNormSoles = 101232
    gCredRefRefCapIntDMMorSoles = 101233
    gCredRefRefCapIntDMVencSoles = 101234
    gCredRefRefCapIntIngDMSoles = 101235
    
    gCredRefCapIntNormDolares = 101254
    gCredRefCapIntMorDolares = 101255
    gCredRefCapIntVencDolares = 101256
    gCredRefCapIntIngDolares = 101267
    
    gCredRefRefCapIntNormDolares = 101257
    gCredRefRefCapIntMorDolares = 101258
    gCredRefRefCapIntVencDolares = 101259
    gCredRefRefCapIntIngDolares = 101260
    
    gCredRefCapIntDMNormDolares = 101236
    gCredRefCapIntDMMorDolares = 101237
    gCredRefCapIntDMVencDolares = 101238
    gCredRefCapIntIngDMDolares = 101268
    
    gCredRefRefCapIntDMNormDolares = 101239
    gCredRefRefCapIntDMMorDolares = 101240
    gCredRefRefCapIntDMVencDolares = 101241
    gCredRefRefCapIntIngDMDolares = 101242
    
    gCredRefIntSuspenso = 101243
    gCredRefGastosSuspenso = 101244
    gCredRefRevIntSuspenso = 101245
    gCredRefRevGastosSuspenso = 101246
    
    
    gCredMiVivCondonacion = 101301
    
    gCredPasoARecup = 130100
    
    'EXTORNOS
    gCredExtPago = 109001
    gCredExtDesemb = 109002
    gCredExtPagoLote = 109003
    gCredExtAprobacion = 109004 'por desaparecer
    gCredExtPagoNormalLote = 109008 'RIRO20141025 ERS122
    gCredExtVigenciaRefinanciado = 109010 'LUCV20160601, ERS004 - 2016
    
    'Llamada de Otra CMAC 1070
    geCredPagodeOtCj = 107100
    'geCredCanceldeOtCj = 127400
    
    'Recepcion en Otra Cmac
    geCredPagoEnOtCj = 126400
    
    'Garantias
    gGarantLevanta = 160101
    gGarantExtorno = 160901
    
    'ITF
    gCredITF = 990103
    gCredITFDesemb = 990106
    
    'ITF en Otra Cmac
    gCredITFEOC = 990303
    
    'RFA
    gColocOpeRFA = 100209
    gColoOpeExRFA = 109005
    
    gColocRFACapital = 1010
    gColocRFAInteresMoratorio = 1108
    gColocRFAInteresCompesatorio = 1107
    gColocRFAInteresDiferido = 1109
    gColocRFAInteresDiferidoRFA = 1110
    
    'Constante de Creditos Automaticos
    gCredAutomatico = 3
    
    'Constante de Estado de Tasacion del Inmueble
    gCredGarantEstadoTasacion = 9061
    
    'Constante de Operacion de Ajuste VAC
    gCredAjusteDiferenciaVAC = 999990

    'Movimiento Operacion de aprobacion de Credito
    gCredAprobacionMov = 109101

    'Movimiento Operacion de aprobacion de Credito
    gCredRetiroMov = 109102
    
    gCredOpeDevJoyaGarantia = 121801 'Devol Joyas - Garantia
    
    'Cambio de Linea de Credito
    '174000 CAMBIO DE LINEA DE CREDITOS NORMALES
    gCredCambioLineaVigNor = 107401
    gCredCambioLineaVigVen = 107402
    gCredCambioLineaRefNor = 107403
    gCredCambioLineaRefVen = 107404
    
    '173000 CAMBIO DE LINEA DE CREDITO EN RECUPERACIONES
    gCredCambioLineaRecJud = 107405
    gCredCambioLineaRecCas = 107406
    
    'Condonación de Crédito
    gCredCondonacionVigNor = 107601
    gCredCondonacionVigVen = 107602
    gCredCondonacionVigMor = 107607
    gCredCondonacionRefNor = 107603
    gCredCondonacionRefVen = 107604
    gCredCondonacionRefMor = 107608
    gCredCondonacionRecJud = 107605
    gCredCondonacionRecCas = 107606
    
    gCredPagoPoliza = 100208 'ARCV 12-03-2007
    
    'JUEZ 20130429 *******************
    gAsientoGarantAvales = 107131
    gReversionGarantAvales = 107331
    gAsientoCondonacionRecup = 107141
    'END JUEZ ************************
    gReversionReprogRoyaAmarilla = 107332 'JUEZ 20131022
    'JUEZ 20140122 *******************
    gAsientoCapHipoPrimeraVivienda = 107133
    gAsientoCapHipoNoPrimeraVivienda = 107134
    gReversionCapHipoPrimeraVivienda = 107333
    gReversionCapHipoNoPrimeraVivienda = 107334
    'END JUEZ ************************
    gReversionReprogDesNat = 107335 'EJVG20140325
    gReversionReprogZonaInundada = "107336" 'PASI 20170405
    'Cajero Corresponsal 20160323
    gPagoCuotaCajeroCorresponsal = 100810
    gCredPagRefMorEfec_CC = 100811
    gCredPagRefNorEfec_CC = 100812
    gCredPagRefVenEfec_CC = 100813
    gCredPagNorMorEfec_CC = 100814
    gCredPagNorNorEfec_CC = 100815
    gCredPagNorVenEfec_CC = 100816

End Enum

Public Enum gCredGarantEstTasacion
    'Estado de Tasacion de Inmueble
    gCredGarantEstTasacionLevantado = 1
    gcredGarantEstTasacionGravado = 2
End Enum
'Caja Chica Desembolso

Public Enum gDatosCajaChica
     gMontoAsig = 1
     gMontoDesemb = 2
     gMontoTope = 3
     gSaldoActual = 4
     gSaldoAnterior = 5
     gNroCajaChica = 6
End Enum

Public Enum gOpeAutorizacion
     gOpeAutorizacionNotaCargoAbono = 999991 'FRHU 20150127 ERS048-2014
     gOpeAutorizacionCancelacion = 999997
     gOpeAutorizacionRetiro = 999996
     gOpeAutorizacionCargoCuenta = 999998 'JUEZ 20131211
     gOpeAutorizacionOtrosEgresosEfectivo = 999999 'FRHU 20140505 ERS063-2014
     gOpeAutorizacion04SueldoCTS = 999992 'APRI20170601 ERS033-2017
End Enum

'**DAOR 20070131, Parametros para generar el archivo de
'**desembolsos en el banco de la nación
Public Enum ParamDesemBcoNac
    gDesBcoNacCodCliCMACMaynas = 312
    gDesBcoNacCtaMN = 313
    gDesBcoNacCodigoBanco = 314
    gDesBcoNacCodigoPrestamo = 315
    
    'Tipos de registro
    gDesBcoNacTipRegIngresoDesembolsos = 316
    gDesBcoNacTipRegAnulacionDesembolsos = 317
    gDesBcoNacTipRegErradosIngDesembolsos = 318
    gDesBcoNacTipRegErradosAnulaciones = 319
    gDesBcoNacTipRegBloqueoDesembolsos = 320
    gDesBcoNacTipRegDesbloqueoDesembolsos = 321
    
    'Tipos de datos
    gDesBcoNacTipDatIngreDesemEfectivo = 322
    gDesBcoNacTipDatIngreDesemChequeGerencia = 323
    gDesBcoNacTipDatDesemTransfMBanco = 324
    gDesBcoNacTipDatDesemTransfOBanco = 325
    gDesBcoNacTipDatAnulacDesemEfecChequeGer = 326
    
    'Tipos de Registro de Recuperaciones
    gRecBcoNacTipRegRecupIngresoRecup = 329
    
    'Tipo de Situaciones del Contrato
    gRecBcoNacTipSitContPendienteRecup = 330
    gRecBcoNacTipSitContAnulado = 331
    
    gRecBcoNacCtaMN = 332
    'By Capi 10032008
    gCapCtaWesternLavDineroME = 332
    gCapCtaWesternLavDineroMN = 333
    
End Enum

'**DAOR 20070131, Estados del Credito en Generación de Archivo
'**para desembolsos en el Banco de la Nación, sólo para aquellos creditos
'**que serán desembolsados en el banco de la nación
Public Enum EstadoCredDesemBcoNac
    gDesBcoNacNuevo = 0 'Nuevo, por generar
    gDesBcoNacGenerado = 1 'Generado
End Enum

'********************************************************
'********************************************************

'Constantes de Impresion a WORD

'cPlantillaCartaAMoroso1 = "\FormatoCarta\CartaMorosoAvisoPago.doc"
'cPlantillaCartaAMoroso2 = "\FormatoCarta\CartaMorosoRequerimientoPago.doc"
'cPlantillaCartaAMoroso3 = "\FormatoCarta\CartaMorosoNotificacionPrejudicial.doc"
'cPlantillaCartaAMoroso4 = "\FormatoCarta\CartaMorosoUltimaNotificacionPrejudicial.doc"
'cPlantillaCartaAMoroso5 = "\FormatoCarta\CartaMorosoCartaNotarial.doc"
'cPlantillaCartaInvCredParalelo = "\FormatoCarta\CartaInvCredParalelo.doc"
'cPlantillaCartaRecup = "\FormatoCarta\CartaRecup.doc"
'EJVG20120328
Public Enum MovLavDineroRelacion
    RealizaTransaccion = 1
    OrdenaTransaccion = 2
    BeneficiaTransaccion = 3
End Enum
Public Type PersonaLavado
    PersCod As String
    Nombre As String
    DocumentoId As String
    Direccion As String
    Ocupacion As String
    'WIOR 20130301 **************
    Nacionalidad As String
    Residente As String
    Peps As String
    'WIOR FIN *******************
End Type
Public Type Recaudo
    CtaCodCredito As String
    CtaCodAbono As String
    PersCod As String
    Placa As String
    EESSFechaHora As Date
    EESSId As String
    EESSNombre As String
    EESSNroTicket As String
    EESSRecaudoBruto As Currency
    EESSITFValorEntrada As Currency
    EESSITFValorSalida As Currency
    EESSRecaudoNeto As Currency
    COFIDEPorcentajeComision As Currency
    IFICod As String
    IFIRecaudoNeto As Currency
End Type
Public Type AbonoEcoTaxi
    CtaCodAbono As String
    Abono As Currency
End Type
Public Enum Alineamiento
    izquierda = 0
    Derecha = 1
    Centro = 2
End Enum
'*** BRGO 20120608 ********
Public Type HabilitaVehiculo
    CtaCod As String
    PersCod As String
    PersNombre As String
    Placa As String
    FechaHora As Date
    IFICod As Integer
    TallerCod As String
End Type
'*** END BRGO
'***Agregado por ELRO el 20120723, según OYP-RFC077-2012
Public Enum TipoCobertura
    gCobBovEfe = 1  'Cobertura de Bóveda para Efectivo
    gCobVenEfe = 2  'Cobertura de Ventanilla para Efectivo
    gCobBovOro = 3  'Cobertura de Bóveda para Oro
    gCobVenOro = 4  'Cobertura de Ventanilla para Oro
End Enum
'***Fin Agregado por ELRO el 20120723*******************
'WIOR 20121015 *****************************************
Public Enum TipoOperacionPersRealiza
    gPersRealizaRetiroEfec = 1  'Retiros en efectivo de titulares personas juridicas
    gPersRealizaRetiroOrdPag = 2  'Retiros con Ordenes de Pago
    gPersRealizaGiros = 3  'Aperturas de Giros
    gPersRealizaGirosCanc = 4  'Cancelaciones de Giros
End Enum
'WIOR FIN **********************************************
'WIOR 20130301 **************
Public Type PersOpeGeneral
    PersCod As String
    TpoCliente As Integer
    TpoDOI As Integer
    NroDOI As String
    Nombre As String
    TipoRelacion As Integer
End Type
 'WIOR FIN *******************

'***************************** Agregado por RIRO el 20130314
Public Enum ConvenioRecaudos
    
    TipoCobro_porConcepto = 0
    TipoCobro_porOperacion = 1
    TipoCalculo_fijo = 0
    TipoCalculo_porcentaje = 1
    Distribucion_fijo = 0
    Distribucion_porcentaje = 1
    
    Convenio_VC = 1
    Convenio_VI = 2
    Convenio_SV = 3
    Convenio_VP = 4
    Convenio_VCM = 5 'RIRO20150120
      
    Registrado = 1
    Pagando = 2
    Pagado = 3
    Anulado = 4
    
    NuevoConvenio = 1
    MantenimientoConvenio = 2
    ConsultaConvenio = 3

    NoTieneWS = 0 'CTI1 ERS027-2019
    SiTieneWS = 1 'CTI1 ERS027-2019
End Enum

Public Enum OperacionesRecaudo

    gDepositoRecaudo = 300120          ' Montos depositadoS
    gDepositoRecaudoCargoCta = 300121  'JUEZ 20150130 Montos depositados Cargo a Cuenta
    gDepositoRecaudoDebitoAuto = 300125 'JUEZ 20150130 Cabecera Pago Debito Automático
    gCargoEmpresaRecaudo = 300111      ' Comision de la empresa
    gCargoClienteRecaudo = 300112      ' Comidion del cliente
    gCargoClienteRecaudoCargoCta = 300113      'JUEZ 20150130 Comision del cliente Cargo a Cuenta
    gIFTRecaudo = 990305               ' Cobro de ITF Recaudo
    gExtornoDepositoRecaudo = 350401   ' Extorno
    gExtornoComisionEmpresa = 350402
    gExtornoComisionCliente = 350403
    gExtornoITF = 350404
    gExtornoServCobRegDebitoAuto = 350405 'JUEZ 20150130

End Enum

'***************************** Fin RIRO
'JUEZ 20150130 *****************************
Public Enum TipoPagoRecaudo
    gPagoNormal = 0
    gPagoDebitoAutomatico = 1
End Enum
'END JUEZ **********************************

'WIOR 20131010 *******************************
Public Type MoraBase
    Usuario As String
    Codigo As String
    Nombre As String
    
    Fecha As Date
    FecMoraBase As Date
    PorcMoraBase As Double
    SaldoHer As Double
    MoraHer As Double
    SaldoSal As Double
    MoraSal As Double
    SaldoPro As Double
    MoraPro As Double
    Saldo As Double
    Mora As Double
    PorcMora As Double
    
    FechaDicAnt As Date
    FecMoraBaseDicAnt As Date
    PorcMoraBaseDicAnt As Double
    SaldoHerDicAnt As Double
    MoraHerDicAnt As Double
    SaldoSalDicAnt As Double
    MoraSalDicAnt As Double
    SaldoProDicAnt As Double
    MoraProDicAnt As Double
    SaldoDicAnt As Double
    MoraDicAnt As Double
    PorcMoraDicAnt As Double
    
    FechaEne As Date
    FecMoraBaseEne As Date
    PorcMoraBaseEne As Double
    SaldoHerEne As Double
    MoraHerEne As Double
    SaldoSalEne As Double
    MoraSalEne As Double
    SaldoProEne As Double
    MoraProEne As Double
    SaldoEne As Double
    MoraEne As Double
    PorcMoraEne As Double
    
    FechaFeb As Date
    FecMoraBaseFeb As Date
    PorcMoraBaseFeb As Double
    SaldoHerFeb As Double
    MoraHerFeb As Double
    SaldoSalFeb As Double
    MoraSalFeb As Double
    SaldoProFeb As Double
    MoraProFeb As Double
    SaldoFeb As Double
    MoraFeb As Double
    PorcMoraFeb As Double
    
    FechaMar As Date
    FecMoraBaseMar As Date
    PorcMoraBaseMar As Double
    SaldoHerMar As Double
    MoraHerMar As Double
    SaldoSalMar As Double
    MoraSalMar As Double
    SaldoProMar As Double
    MoraProMar As Double
    SaldoMar As Double
    MoraMar As Double
    PorcMoraMar As Double
    
    FechaAbr As Date
    FecMoraBaseAbr As Date
    PorcMoraBaseAbr As Double
    SaldoHerAbr As Double
    MoraHerAbr As Double
    SaldoSalAbr As Double
    MoraSalAbr As Double
    SaldoProAbr As Double
    MoraProAbr As Double
    SaldoAbr As Double
    MoraAbr As Double
    PorcMoraAbr As Double
    
    FechaMay As Date
    FecMoraBaseMay As Date
    PorcMoraBaseMay As Double
    SaldoHerMay As Double
    MoraHerMay As Double
    SaldoSalMay As Double
    MoraSalMay As Double
    SaldoProMay As Double
    MoraProMay As Double
    SaldoMay As Double
    MoraMay As Double
    PorcMoraMay As Double
    
    FechaJun As Date
    FecMoraBaseJun As Date
    PorcMoraBaseJun As Double
    SaldoHerJun As Double
    MoraHerJun As Double
    SaldoSalJun As Double
    MoraSalJun As Double
    SaldoProJun As Double
    MoraProJun As Double
    SaldoJun As Double
    MoraJun As Double
    PorcMoraJun As Double
    
    FechaJul As Date
    FecMoraBaseJul As Date
    PorcMoraBaseJul As Double
    SaldoHerJul As Double
    MoraHerJul As Double
    SaldoSalJul As Double
    MoraSalJul As Double
    SaldoProJul As Double
    MoraProJul As Double
    SaldoJul As Double
    MoraJul As Double
    PorcMoraJul As Double
    
    FechaAgo As Date
    FecMoraBaseAgo As Date
    PorcMoraBaseAgo As Double
    SaldoHerAgo As Double
    MoraHerAgo As Double
    SaldoSalAgo As Double
    MoraSalAgo As Double
    SaldoProAgo As Double
    MoraProAgo As Double
    SaldoAgo As Double
    MoraAgo As Double
    PorcMoraAgo As Double
    
    FechaSep As Date
    FecMoraBaseSep As Date
    PorcMoraBaseSep As Double
    SaldoHerSep As Double
    MoraHerSep As Double
    SaldoSalSep As Double
    MoraSalSep As Double
    SaldoProSep As Double
    MoraProSep As Double
    SaldoSep As Double
    MoraSep As Double
    PorcMoraSep As Double
    
    FechaOct As Date
    FecMoraBaseOct As Date
    PorcMoraBaseOct As Double
    SaldoHerOct As Double
    MoraHerOct As Double
    SaldoSalOct As Double
    MoraSalOct As Double
    SaldoProOct As Double
    MoraProOct As Double
    SaldoOct As Double
    MoraOct As Double
    PorcMoraOct As Double
    
    FechaNov As Date
    FecMoraBaseNov As Date
    PorcMoraBaseNov As Double
    SaldoHerNov As Double
    MoraHerNov As Double
    SaldoSalNov As Double
    MoraSalNov As Double
    SaldoProNov As Double
    MoraProNov As Double
    SaldoNov As Double
    MoraNov As Double
    PorcMoraNov As Double
    
    FechaDic As Date
    FecMoraBaseDic As Date
    PorcMoraBaseDic As Double
    SaldoHerDic As Double
    MoraHerDic As Double
    SaldoSalDic As Double
    MoraSalDic As Double
    SaldoProDic As Double
    MoraProDic As Double
    SaldoDic As Double
    MoraDic As Double
    PorcMoraDic As Double

End Type

Public Type AnalistaBPP
    cPersCod As String
    CodAge As String
    Agencia As String
    Usuario  As String
    NombreAnalista  As String
    Nivel As Integer
    Comite As Integer
    Categoria As String
    MoraBase As Double
    '*************METAS DE CUMPLIMIENTO***********
        'CRECIMIENTO DE SALDO
        MetaSaldo As Double
        MetaSaldoAG As Double
        'CRECIMIENTO DE CLIENTE
        MetaCliente As Integer
        MetaClienteAG As Integer
        'NÚMERO DE OPERACIONES
        MetaOperaciones As Integer
        MetaOperacionesAG As Integer
        'MORA ACEPTABLE DE 8 A 30 DIAS
        MetaMora  As Double
        MetaMoraAG As Double
        'META RENDIMIENTO
        MetaRendimiento As Double
        MetaRendimientoAG As Double
    '*********************************************
    'CRECIMIENTO DE SALDO
    SaldoCapital As Double
    SaldoInicial As Double
    SaldoEntrante As Double
    SaldoSaliente As Double
    SIA As Double
    SCE As Double
    PPOSaldo As Double
    PTFSaldo As Double
    IXSaldo As Double
    PXSaldo As Double
    'CRECIMIENTO DE CLIENTE
    CantClientes As Integer
    ClientesInicial As Integer
    ClientesEntrantes As Integer
    ClientesSalientes As Integer
    NIC As Double
    NCE As Double
    PPOCliente As Double
    PTFCliente As Double
    IXCliente As Double
    PXCliente As Double
    'NÚMERO DE OPERACIONES
    NFO1 As Double
    NOE1 As Double
    PPOOpe1 As Double
    NFO2 As Double
    NOE2 As Double
    PPOOpe2 As Double
    PTFO As Double
    IXOperaciones As Double
    PXOperaciones As Double
    'MORA ACEPTABLE 8-30
    MF830 As Double
    MI830 As Double
    ME830 As Double
    PP830 As Double
    PTFMora As Double
    IXM830 As Double
    PXMora As Double
    'RENDIMIENTO DE CARTERA
    ICOB As Double
    PESP As Double
    CCC As Double
    RCA As Double
    IXRendimiento As Double
    'MORA >30 DÍAS
    MIMayor30 As Double
    CJI As Double
    TMI As Double
    MFMayor30 As Double
    CJF As Double
    TMF As Double
    'RESUMEN DEL BONO
    BonoMeta As Double
    BonoPlus As Double
    BonoRendimiento As Double
    Penalidad As Double
    BonoTotal As Double
    '*****************DATOS ADICIONALES*************
        'CRECIMIENTO DE SALDO
        TopeSaldo As Double
        PorcMinSaldo As Double
        'CRECIMIENTO DE CLIENTES
        TopeCliente As Double
        PorcMinCliente As Double
        'NÚMEREO DE OPERACIONES
        TopeOperaciones As Double
        PorcMinOperaciones As Double
        'MORA 8-30
        RangoPerMora As Double
        TopeMora As Double
        PorcMinMora As Double
        'MORA >30
        MoraAcepMayor30 As Double
        'RENDIMIENTO CMACM
        IntCobCMACM As Double
        SaldoCMACM As Double
        RendCMACM As Double
        MinRendCartera As Double
        'FACTOR DE RENDIMIENTO DE CARTERA
        FactorRend As Double
        'BONO PLUS Y %OPERACIONES
        PrimQuincena As Double
        SegunQuincena As Double
        SaldoPlus As Double
        ClientesPlus As Double
        OperacionesPlus As Double
        MoraPlus As Double
        'MORAS
        Mora830IncialCierre  As Double
        Mora830Entrante As Double
        Mora830Saliente As Double
    '*********************************************
End Type

Public Type CoordinadorJABPP
        cCodAge As String
        Agencia As String
        nCargo As Integer
        Cargo As String
        cPersCod As String
        Comite As String
        Usuario As String
        Nombre As String
        CantAnalista As Integer
        PorBonificacion As Double
        SaldoCartera As Double
        SaldoCarteraCierre As Double
        PorSalCartVencJud As Double
        PorcSalVenJud As Double
        RendCaja As Double
        RendAG As Double
        Tope As Double
        BonoTotal As Double
        PorBonificacionConf As Double
        AnalistaBoni As Integer
        SaldoCapital As Double
        SaldoVencJud As Double
        IntCob As Double
End Type

Public Type JefeTerritorialBPP
        nZona As String
        Zona As String
        cPersCod As String
        Usuario As String
        Nombre As String
        SaldoCartera As Double
        SaldoCarteraCierre As Double
        PorSalCartVencJud As Double
        PorcSalVenJud As Double
        RendCaja As Double
        Rend As Double
        CantAge As Integer
        AgeBoni As Integer
        PorBoni As Double
        PorConfJF As Double
        Tope As Double
        BonoTotal As Double
End Type
'WIOR FIN ************************************
'EJVG20140131 ***
Public Enum TipoOperacionCheque
    Ninguno = 0 'Ninguno
    DPF_Apertura = 1 'Apertura DPF
    DPF_AumentoCapital = 2 'Aumento Capital DPF
    AHO_Deposito = 3 'Deposito Ahorro
    CTS_Deposito = 4 'Deposito CTS
    CTS_DepositoLote = 5 'Deposito en Lote CTS
    CRED_Pago = 6 'Pago Creditos
    CRED_PagoLote = 7 'Pago en Lote Creditos
    AHO_DepositoLote = 8 'Deposito Ahorro Lote
    CRED_PagoNormalLote = 9 ' RIRO20141014 ERS122 Pago Normal en Lote de Creditos
    CRED_LiqCreditosSegDes = 10 'WIOR 20160426 Liquidacion Credito con el Seguro Desgravamen
    AHO_Apertura = 11 'Aperura Ahorro
    AHO_AperturaLote = 12 'Apertura Lote Ahorro
    AHO_DepositoHaberesLote = 13 'Deposito Haberes Lote
    CTS_Apertura = 14 ' Apertura CTS
    DPF_AperturaLote = 15 'Apertura Lote CTS
End Enum
'END EJVG *******
'ALPA20140421*******************************
Public Enum ConstanteNivelRiesgo
    gConsNivelRiesgos = 9999
    gConsNivelRiesgosCobNR1 = 9028
    gConsNivelRiesgosCobNR2 = 9029
End Enum
'*******************************************
'FRHU 20140505 ERS063-2014
Public Enum OtrasOperacionesAutorizaEstado
    gOtraOperacEstSolicitado = 0
    gOtraOperacEstAprobado = 1
    gOtraOperacEstRechazado = 2
End Enum
'FIN FRHU 20140505
'FRHU 20140519 ERS068-2014
Public Enum EstadoDeActivacionDeSeguroDeTarjeta
    gSegTarjetaSolicitud = 1000
    gSegTarjetaAprobado = 1001
    gSegTarjetaAceptado = 1002
    gSegTarjetaRechazado = 1003
End Enum
'FIN FRHU 20140519
'WIOR 2014606 ******************************
Public Type PolizaRenova
        cCtaCod As String
        cNumGarant As String
        cNumPoliza As String
        nTipoRenovacion As Integer
        cPersCodAseguradora As String
        cNumCertificado As String
        nMoneda As Integer
        dTasacion As Date
        nValorEdificacion As Double
        nInmueble As Integer
        nPrimaNeta As Double
        nDerechoEm As Double
        nIGV As Double
        nPrimaTotal As Double
        nPrimaTotalTC As Double
        nDesde As Integer
        nPrimaCuota As Double
        dVigencia As Date
        dVencimiento As Date
        dVigenciaNueva As Date
        dVencimientoNueva As Date
        cMovRegistro As String
        cAgeCod As String
        cDescripcion As String
        nMonedaAux As Integer
        cDesInmueble As String
End Type

Public Type CredPagoAuto
        cCtaCod As String
        cTpoCredCod As String
        nDiasAtraso As Integer
        nNroCalen As Integer
        CuotaMin As Integer
        CuotaMax As Integer
        Pagar As Double
        cMetLiquidacion As String
        AhorroDebitar As Variant
End Type
'WIOR FIN ************************************
'JUEZ 20141212 *********************************
Public Enum CredTiposPagosNormales
    gPagoNorMaxDosCuotasPorVenc = 1
    gPagoVencSaldoPend = 2
    gPagoVencSaldoPorVenc = 3
    gPagoNorEnFechaCuotaSaldoPend = 4
    gPagoNorAntesFechaCuotaSaldoPorVenc = 5
End Enum
'END JUEZ **************************************
'RECO20150310 ERS149-2014**********************
Public Enum TipoSeguro
    gDesgravament = 10010000
    gContraIncendio = 10020000
    gMYPE = 10030000
    gSepelio = 10040000
End Enum
Public Enum SolicitudCoberEstado
    gSolicitudCoberPendiente = 1
    gSolicitudCoberRechazada = 2
    gSolicitudCoberReconsideracion = 3
    gSolicitudCoberAceptada = 4
    gSolicitudCoberExtorno = 5
End Enum
'RECO FIN**************************************
'JUEZ 20150130 *********************************
Public Enum CaptacServDebitoAutomaticoTipo
    gServConvenio = 1
    gServCredito = 2
End Enum
'END JUEZ **************************************
'JUEZ 20150415 **************************
Public Enum CredTiposPagosEspecial
    gAdelantoCuota = 1
    gPagoAnticipado = 2
End Enum
'END JUEZ *******************************

'RIRO 20150701 ERS162-2014 *********************
Public Type PagoUtilidades

    nIdUtilidad As Long
    nIdTrama As Long
    nPeriodo As Integer
    sNombre As String
    sDOI As String
    sCargo As String
    sArea As String
    dFechaIngreso As Date
    nImporte As Double
    sMoneda As String
    sCiudad As String
    
    C09_ParticipAdistribuir As Double
    C10_DiasLaborTodosTrabAnio As Double
    C11_RemunPercibTodosTrabAnio As Double
    C12_MontDistribXdiasLabor As Double
    C13_MontDistribXremunPercib As Double
    C14_TotDiasEfectivLabor As Integer
    C15_TotRemuneraciones As Double
    C16_ParticipXdiasLabor As Double
    C17_ParticipXremuneraciones As Double
    C18_TotParticipUtilidades As Double
    C19_RetencionImpuestoRenta As Double
    C20_TotalDescuento As Double
    C21_TotalPagar As Double
    
End Type
'END RIRO **************************************
'FRHU 20151130 ERS077-2015
Public Type TActAutDatos
    sUbicGeografica As String
    sNombres As String
    sApePat As String
    sApeMat As String
    sApeCas As String
    sSexo As String
    sEstadoCivil As String
    sDomicilio As String
    cNacionalidad As String
    sRefDomicilio As String
    sTelefonos As String
    sCelular As String
    sEmail As String
    sPersIDTpo As String
    sPersIDnro As String
End Type
'FIN FRHU 20151130
'RECO20151123 ERS073-2015 **************************
Public Enum SegCertificadoTipo
    gSegCertificadoTarjeta = 1
    gSegCertificadoSepelio = 2
End Enum
Public Enum EstadoSegSepelio
    gSegEstadoAfiliado = 1
    gSegEstadoDesafiliado = 2
    gSegEstadoAnulado = 3
    gSegEstadoExtorno = 4
End Enum
Public Enum TipoBusquedaSepelio
    gSegTpoBusCta = 1
    gSegTpoBusPersCod = 2
    gSegTpoBusNumTarj = 3
End Enum
'RECO FIN ******************************************
'RECO20160208 ERS056-2015***************************
Public Enum PigTipoTarifarioNotif
    gPigTarifarioAge = 1
    gPigTarifarioMinka = 2
End Enum
'RECO FIN*******************************************
'EJVG20150112 ***
Public Enum eTipoInicioGarantia
    RegistrarGarantia = 1
    EditarGarantia = 2
    ConsultarGarantia = 3
End Enum
Public Enum eGarantiaCambiosFila
    GarantFilaSinCambios = 0
    GarantFilaNueva = 1
    GarantFilaModificada = 2
    GarantFilaEliminada = 3
    GarantFilaOculta = 4
End Enum
Public Enum eGarantiaClasificacionBien
    GarantiaBienInmuebles = 1
    GarantiaBienMuebles = 2
    GarantiaBienVehiculos = 3
    GarantiaBienTitulosValores = 4
    GarantiaBienCoberturasPersonales = 5
End Enum
Public Enum eGarantiaPropietarioRelacion
    GarantiaPropietarioRelacionTitular = 1
    GarantiaPropietarioRelacionConyugue = 2
    GarantiaPropietarioRelacionRepresentante = 3
    GarantiaPropietarioRelacionOtros = 4
End Enum
Public Enum eGarantiaTipoValorizacion
    ValorDirectoTotal = 1
    ValorDirectoDetallado = 2
    GarantiaAutoliquidable = 3
    TasacionVehicular = 4
    TasacionMetalesPreciosos = 5
    TasacionInmobiliaria = 6
    OtrasTasacionesMobiliarias = 7
    GravamenFavorOtraIFi = 8 'EJVG20160224
End Enum
Public Enum eTipoRelacion
    Titular = 1
    Conyugue = 2
    Representante = 3
    Otros = 4
End Enum
Public Enum eGarantiaTramiteLegalTipo
    BloqueoRegistral = 1
    PrimeraPreferente = 2
    SegundaPreferente = 3
    AmpliacionGarantia = 4
    ModificacionGarantia = 5
    LevantamientoGarantia = 6
End Enum
Public Enum eGarantiaTramiteLegalEstado
    Pendiente = 1
    Inscrita = 2
End Enum
Public Enum eGarantiaTipoCobertura
    Sabana = 1
    Especifica = 2
End Enum
Public Enum eGarantiaTipoRetiroRapiFlash
    nRetiroIntAdelantado = 4
    nRetiroIntMensual = 1
    nRetiroIntFinal = 2
End Enum
Public Enum eTipoInicioPolizaExt
    PolizaExtxSolicitud = 1
    PolizaExtxRegistroManual = 2
    PolizaExtxRegistroCreditoVigente = 3
    PolizaExtxSustitucion = 4
    PolizaExtxRenovacion = 5
End Enum
Public Enum eTipoInicioPolizaInt
    PolizaIntxSolicitud = 1
    PolizaIntxRegistroManual = 2
    PolizaIntxRegistroCreditoVigente = 3
    PolizaIntxSustitucion = 4
End Enum
Public Enum ePolizaTipo
    Incendio = 1
    Vehicular = 2
    Cascos = 3
    MaquinariaEquipo = 4
    Mercaderia = 5
End Enum
Public Type tPropietarioRelacion
    sPersCod As String
    sPersNombre As String
    nRelacionTpo As eTipoRelacion
    sRelacionTpo As String
    nGarantiaCambiosFila As eGarantiaCambiosFila
End Type
Public Type tValorDirectoTotal
    nVRM As Currency
End Type
Public Type tValorDirectoDetallado
    sDescripcion As String
    nCantidad As Integer
    nValor As Currency
    nDocTpo As Integer
    sDocTpo As String
    sDocNro As String
End Type
Public Type tValorAutoLiquidable
    sCtaCod As String
    nFormaRetiro As Integer
    sFormaRetiro As String
    nMontoCuenta As Currency
    nMontoCuentaDisp As Currency
    nMontoDisponible As Currency
    nMontoCuentaNEW As Currency
    nMontoCuentaDispNEW As Currency
    nMontoDisponibleNEW As Currency
End Type
Public Type tValorTasacionInmobiliaria
    sUbicacionGeografica As String
    sDireccion As String
    nValorComercial As Currency
    nValorEdificacion As Currency
    nVRM As Currency
    nClase As Integer
    nCategoria As Integer
    nNroLocales As Integer
    nNroPisos As Integer
    nNroSotanos As Integer
    nAnioConstruccion As Integer
    dTasacion As Date
    nTasacionTC As Currency
    sTasadorCod As String
    sTasadorNombre As String
    sTasadorDNI As String
    sTasadorREPEV As String
    bMigrado As Boolean
    sMz As String 'CTI5 ERS0012020
    sLt As String 'CTI5 ERS0012020
    sEt As String 'CTI5 ERS0012020
    
End Type
Public Type tValorTasacionVehicular
    sDescripcion As String
    sCarroceria As String
    sMarca As String
    sModelo As String
    sNroPlaca As String
    sNroMotor As String
    sNroChasis As String
    nAnioFabricacion As Integer
    nAnioAdquisicionTienda As Integer
    nValorComercial As Currency
    nVRM As Currency
    dTasacion As Date
    nTasacionTC As Currency
    sTasadorCod As String
    sTasadorNombre As String
    sTasadorDNI As String
    sTasadorREPEV As String
    bMigrado As Boolean
End Type
Public Type tValorTasacionMobiliariaOtras
    sDescripcion As String
    sNroSerie As String
    sMarca As String
    sModelo As String
    nAnioFabricacion As Integer
    nValorComercial As Currency
    nVRM As Currency
    dTasacion As Date
    nTasacionTC As Currency
    sTasadorCod As String
    sTasadorNombre As String
    sTasadorDNI As String
    sTasadorREPEV As String
    bMigrado As Boolean
End Type
'EJVG20160224 ***
Public Type tValorGravamenFavorOtraIFiDet
    sIFTpo As String
    sIFICod As String
    sIFINombre As String
    nGravado As Currency
End Type
Public Type tValorGravamenFavorOtraIFi
    sDescripcion As String
    nValorComercial As Currency
    nOpTpRegPub As Integer 'JOEP20180524 ERS009-2018
    vValorGravamenFavorOtraIFiDet() As tValorGravamenFavorOtraIFiDet
    bMigrado As Boolean
End Type
'END EJVG *******
Public Type tValorizacion
    dfecha As Date
    nValorizacionTpo As eGarantiaTipoValorizacion
    sValorizacionTpo As String
    nVRM As Currency
    sGlosa As String
    nVTpRegPub As Integer 'JOEP20180524 ERS009-2018
    nGarantiaCambiosFila As eGarantiaCambiosFila
    bVinculado As Boolean
    vValorDirectoTotal As tValorDirectoTotal
    vValorDirDet() As tValorDirectoDetallado
    vValorAutoLiquidable As tValorAutoLiquidable
    vValorTasacionInmobiliaria As tValorTasacionInmobiliaria
    vValorTasacionVehicular As tValorTasacionVehicular
    vValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras
    vValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi 'EJVG20160224
    sUltimaActualizacion As String
    bEliminar As Boolean
End Type
Public Type tTramiteLegal
    dfecha As Date
    sNroPartidaRegistral As String
    nTipoTramite As eGarantiaTramiteLegalTipo
    cTipoTramite As String
    nGravamen As Currency
    nZonaRegistralID As Integer
    cZonaRegistralNombre As String
    nOficinaRegistralID As Integer
    cOficinaRegistralNombre As String
    sNotariaCod As String
    sNotariaNombre As String
    nEstado As eGarantiaTramiteLegalEstado
    cEstado As String
    nTipoCobertura As eGarantiaTipoCobertura
    sNroAsiento As String
    dInscripcion As Date
    nGarantiaCambiosFila As eGarantiaCambiosFila
    bVinculado As Boolean
    sUltimaActualizacion As String
    bMigrado As Boolean
    bEliminar As Boolean
End Type
Public Type tGarantiaGravamen
    sNumGarant As String
    nMoneda As Moneda
    dFechaValorizacion As Date
    dFechaTramiteLegal As Date
    nSaldoCobertura As Double
    nRatio As Double
    nMaxCobertura As Double
    nCobertura As Double
    nOrden As Integer
End Type
'END EJVG *******
'JUEZ 20160216 *************************
Public Enum ColocacReprogEstado
    gEstReprogSolicitado = 200
    gEstReprogPropuesto = 201
    gEstReprogAutorizado = 202
    gEstReprogAprobado = 203
    gEstReprogRechazado = 204
    gEstReprogReprogramado = 205
    gEstReprogVBAdmCred = 206
    gEstReprogExtornado = 207 'JOEP20171129
    gEstReprogSolicitadoReprogramado = 208 'Agrego JOEP20171214 ACTA220-2017
End Enum
'END JUEZ ******************************

'RIRO20160407 Tarifario Versionado
Public Type tComision
    IdComision As Integer
    Producto As Integer
    SubProducto As Integer
    Personeria As Integer
    Version As Integer
    FechaRegistro As Date
    Glosa As String
    Estado As Integer
    FechaVigencia As Date
    Grupo As String
    MovNro As String
    CtaDepositoSolesTipo As Integer
    CtaDepositoSolesComision As Double
    CtaDepositoSolesMin As Double
    CtaDepositoSolesMax As Double
    CtaDepositoDolaresTipo As Integer
    CtaDepositoDolaresComision As Double
    CtaDepositoDolaresMin As Double
    CtaDepositoDolaresMax As Double
    CtaRetiroSolesTipo As Integer
    CtaRetiroSolesComision As Double
    CtaRetiroSolesMin As Double
    CtaRetiroSolesMax As Double
    CtaRetiroDolaresTipo As Integer
    CtaRetiroDolaresComision As Double
    CtaRetiroDolaresMin As Double
    CtaRetiroDolaresMax As Double
    CtaTransfSolesTipo As Integer
    CtaTransfSolesComision As Double
    CtaTransfSolesMin As Double
    CtaTransfSolesMax As Double
    CtaTransfDolaresTipo As Integer
    CtaTransfDolaresComision As Double
    CtaTransfDolaresMin As Double
    CtaTransfDolaresMax As Double
    CajOperacionesLibresATM As Integer
    CajOperacionesLibresMON As Integer
    CajRetiroSolesExcesoOperacionesATM As Double
    CajRetiroDolaresExcesoOperacionesATM As Double
    CajConsultaSaldosSolesExcesoOperacionesATM As Double
    CajConsultaSaldosDolaresExcesoOperacionesATM As Double
    CajConsultaMovSolesExcesoOperacionesATM As Double
    CajConsultaMovDolaresExcesoOperacionesATM As Double
    CajCambioClaveSolesExcesoOperacionesATM As Double
    CajCambioClaveDolaresExcesoOperacionesATM As Double
    CajRetiroSolesExcesoOperaciones05MON As Double
    CajRetiroSolesExcesoOperaciones10MON As Double
    CajRetiroSolesExcesoOperaciones15MON As Double
    CajRetiroSolesExcesoOperaciones20MON As Double
    CajRetiroSolesExcesoOperaciones25MON As Double
    CajRetiroSolesExcesoOperaciones30MON As Double
    CajRetiroSolesExcesoOperaciones35MON As Double
    CajRetiroSolesExcesoOperaciones40MON As Double
    CajRetiroSolesExcesoOperaciones45MON As Double
    CajRetiroSolesOTRO As Double
    CajRetiroDolaresOTRO As Double
    CajConsultaSaldosSolesOTRO As Double
    CajConsultaSaldosDolaresOTRO As Double
    CajRetiroSolesATMextranjero As Double
    CajRetiroDolaresATMextranjero As Double
    CajConsultaSaldosSolesATMextranjero As Double
    CajConsultaSaldosDolaresATMextranjero As Double
    CajCambioClaveSolesATMextranjero As Double
    CajCambioClaveDolaresATMextranjero As Double
    CajCompraSolesNacionalVISA As Double
    CajCompraDolaresNacionalVISA As Double
    CajCompraSolesInternacionalVISA As Double
    CajCompraDolaresInternacionalVISA As Double
    VntOperacionesLibresRetiro As Integer
    VntOperacionesLibresDeposito As Integer
    VntOperacionesLibresTransf As Integer
    VntOperacionesLibresOrdenesPago As Integer
    VntOperacionesLibresInterplazaRetiro As Integer
    VntOperacionesLibresInterplazaDeposito As Integer
    VntOperacionesLibresInterplazaTransf As Integer
    VntExcesoOperacionesOrdenesPagoSoles As Double
    VntExcesoOperacionesOrdenesPagoDolares As Double
    VntRetiroSolesExcesoOperaciones As Double
    VntRetiroDolaresExcesoOperaciones As Double
    VntDepositoSolesExcesoOperaciones As Double
    VntDepositoDolaresExcesoOperaciones As Double
    VntTranfSolesExcesoOperaciones As Double
    VntTranfDolaresExcesoOperaciones As Double
    VntExtractoCtaAhorrosSoles As Double
    VntExtractoCtaAhorrosDolares As Double
    VntExtractoCtaAhorrosValor As Integer
    VntConsultaSaldosVentanillaSoles As Double
    VntConsultaSaldosVentanillaDolares As Double
    VntConsultaSaldosVentanillaValor As Integer
    VntRetiroSinTarjetaDebitoSoles As Double
    VntRetiroSinTarjetaDebitoDolares As Double
    VntRetiroSinTarjetaDebitoValor As Integer
    SrvMantenimientoCtaSoles As Double
    SrvMantenimientoCtaDolares As Double
    SrvMantenimientoCtaValor As Integer
    SrvDebitoServicioRecaudoSoles As Double
    SrvDebitoServicioRecaudoDolares As Double
    SrvDebitoServicioRecaudoValor As Integer
    SrvDebitoPagoCreditoSoles As Double
    SrvDebitoPagoCreditoDolares As Double
    SrvDebitoPagoCreditoValor As Integer
    SrvEnvioEstadoCtaFisicoSoles As Double
    SrvEnvioEstadoCtaFisicoDolares As Double
    SrvEnvioEstadoCtaFisicoValor As Integer
    SrvReposicionTarjetaDebitoSoles As Double
    SrvReposicionTarjetaDebitoDolares As Double
    SrvReposicionTarjetaDebitoValor As Integer
End Type
Public Type tTasa
    IdTasa As Integer
    Producto As Integer
    SubProducto As Integer
    Personeria As Integer
    Grupo As String
    Version As Integer
    MovNro As String
    FechaRegistro As Date
    Glosa As String
    Estado As Integer
    AplicaInmediato As Integer
    rsTasa As ADODB.Recordset
End Type
'End Tarifario Versionado
'EJVG20160510 ***
Public Enum eProcesoEmiteInformeRiesgo
    Sugerencia = 1
    NivelAprobacion = 2
    Aprobacion = 3
End Enum
'END EJVG *******

'-- PEAC 20160618
'Public Type tForEvalEstFinFormato6 'LUCV20171015 Comentó, según ERS0512017
Public Type tFormEvalDetalleEstFinFormato6
'    nConsCod As Integer
'    nConsValor As Integer
    dfecha As Date
    cDescripcion As String
    nImporte As Double
End Type

'-- PEAC 20160621
'Public Type tForEvalResumenEstFinFormato6 'LUCV20151015 Comentó, según ERS0512017
Public Type tFormEvalPrincipalEstFinFormato6
    cConcepto As String
    nImportePP As Double
    nImportePE As Double
    dfecha As Date
    nConsCod As Integer
    nConsValor As Integer
    vTotal() As tFormEvalDetalleEstFinFormato6 'LUCV20171510 Agregó, según ERS0512017
    nTotal As Double 'LUCV20171510 Agregó, según ERS0512017
    vPP() As tFormEvalDetalleEstFinFormato6 'tForEvalEstFinFormato6 'LUCV20171510, Comentó según ERS0512017
    vPE() As tFormEvalDetalleEstFinFormato6 'tForEvalEstFinFormato6 'LUCV20171510, Comentó según ERS0512017
    nDetPP As Integer
    nDetPE As Integer
End Type

'LUCV20160705, Matrices Type
'DetalleMatrix->**********Activos
Public Type tFormEvalDetalleActivosCtasCobrarFormato5 'Ctas x Cobrar
    dfecha As Date
    cCtaporCobrar As String
    nTotal As Currency
End Type
Public Type tFormEvalDetalleActivosInventarioFormato5 'Inventario
    cMercaderia As String
    nCantidad As Double
    cUnidMed As String
    nCostoUnit As Currency
    nTotal As Currency
End Type
Public Type tFormEvalDetalleActivosActivoFijoFormato5 'Activo Fijo
    cDescripcion As String
    nCantidad As Double
    nPrecio As Currency
    nTotal As Currency
End Type
'DetalleMatrix->**********Pasivos
Public Type tFormEvalDetallePasivosSobreGirosFormato5  'Sobregiros y Prestamos
    cEntidad As String
    cCredito As String
    nTEA As Double
    nCuotas As Integer
    nMontoCuota As Currency
    nCuotasPend As Double
End Type

'LUCV20160706: Matriz Type Principal Activos - Pasivos **********
Public Type tFormEvalPrincipalActivosFormato5 'ACTIVOS
    cConcepto As String
    nImportePP As Double
    nImportePE As Double
    nConsCod As Integer
    nConsValor As Integer
    vPPActivoCtaCobrar() As tFormEvalDetalleActivosCtasCobrarFormato5 'Ctas x Cobrar
    vPEActivoCtaCobrar() As tFormEvalDetalleActivosCtasCobrarFormato5 'Ctas x Cobrar
    
    vPPActivoInventario() As tFormEvalDetalleActivosInventarioFormato5 'Inventario
    vPEActivoInventario() As tFormEvalDetalleActivosInventarioFormato5 'Inventario
    
    vPPActivoActivoFijo() As tFormEvalDetalleActivosActivoFijoFormato5 'Activo Fijo
    vPPEctivoActivoFijo() As tFormEvalDetalleActivosActivoFijoFormato5 'Activo Fijo
    
    nDet1PP As Integer
    nDet1PE As Integer
    
    nDet2PP As Integer
    nDet2PE As Integer
    
    nDet3PP As Integer
    nDet3PE As Integer
End Type

Public Type tFormEvalPrincipalPasivosFormato5 'PASIVOS
    cConcepto As String
    nImportePP As Double
    nImportePE As Double
    nConsCod As Integer
    nConsValor As Integer

    vPPPasivoSobregiro() As tFormEvalDetallePasivosSobreGirosFormato5 'Sobregiros y Prestamos
    vPEPasivoSobregiro() As tFormEvalDetallePasivosSobreGirosFormato5 'Sobregiros y Prestamos
    
    vPPActivoCtaCobrar() As tFormEvalDetalleActivosCtasCobrarFormato5 'Ctas x Cobrar
    vPEActivoCtaCobrar() As tFormEvalDetalleActivosCtasCobrarFormato5 'Ctas x Cobrar
    
    nDet3PP As Integer
    nDet3PE As Integer
    
    nDet4PP As Integer
    nDet4PE As Integer
End Type
'LUCV20160705, Fin

'LUCV20160627, Según ERS004-2016 ->**********
Public Enum RatioIndicador
gRatioCapacidadPago = 1
gRatioEndeudamiento = 2
gRatioIngresoNetoNego = 3
gRatioExcedenteMensual = 4
gRatioLiquidezCte = 5
gRatioRentabilidadPat = 6
gRatioCapaciPagoRDS = 7
End Enum
''LUCV20160627, Fin <-**********

'LUCV20160629, Según ERS004-2016 ->**********
Public Enum FormatoConstante
gFormatoGastosNego = 7022
gFormatoGastosFami = 7023
gFormatoOtrosIngresos = 7024
gFormatoActivos = 7025
gFormatoPasivos = 7026
gFormatoGastosNegoFlujoCaja = 7027 'CTI320200110 ERS003-2020. Agregó
End Enum
''LUCV20160629, Fin <-**********
                                                                                
'LUCV20160629, Según ERS004-2016 ->**********
Public Enum FormatoCodFlex
gCodDeudaLCNUGastoFami = 7
gCodCuotaCmac = 8
gCodCuotaIfiGastoNego = 9
gCodCuotaIfiGastoFami = 5
'gCodCuotaIfiFlujoCaja = 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
gCodCuotaIfiFlujoCaja = 19 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
gCodCuotaIfiNoSupervisadaFlujoCaja = 22 'CTI320200110 ERS003-2020. Agregó
gCodCuotaIfiNoSupervisadaGastoNego = 13 'CTI320200110 ERS003-2020. Agregó
gCodCuotaIfiNoSupervisadaGastoFami = 9  'CTI320200110 ERS003-2020. Agregó

gCodTotalActivo = 1000
gCodTotalPasivo = 1000
gCodTotalPatrimonio = 1001
End Enum
'LUCV20160629, Fin <-**********
                                       
'LUCV20160707, Según ERS004-2016 ->**********
Public Enum TipoPatrimonio
gCodPatrimonioPersonal = 1
gCodPatrimonioEmpresarial = 2
End Enum
'LUCV20160707, Fin <-**********

'EJVG20160129 ***
Public Type TCompraDeuda
    nID As Integer
    sIFICod As String
    sIFINombre As String
    nDestino As Integer
    bRecompra As Boolean
    sNroCredito As String
    dDesembolso As Date
    nMoneda As Moneda
    nMontoDesembolso As Double
    nNroCuotasPactadas As Integer
    nNroCuotasPagadas As Integer
    nSaldoComprar As Double
    nMontoCuota As Double
    bCreditoAnt As Boolean          '**ARLO20180528
    dCancelacion As Date            '**ARLO20180528
    nMontoDesembolsoAnt As Double   '**ARLO20180528
    nMontoCuotaAnt As Double        '**ARLO20180528
End Type
'INICIO EAAS20180801 SEGUN ERS-054-2018 Agua y Saneamiento ***
Public Type TAguaSaneamiento
    nID As Integer
    nDestinoCod As Integer
    sDestinoDescripcion As String
    sNroCredito As String
    nSubDestinoCod As Integer
    sSubDestinoDescripcion As String
    nMontoP As Double
    nMontoS As Double
    nBeneficia As Integer
    nSumaAguaSaneamiento As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM Crédito Verde
End Type
'FIN EAAS20180801 SEGUN ERS-054-2018 Agua y Saneamiento ***
'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM Crédito Verde ***
Public Type TCreditoVerde
    nID As Integer
    nDestinoCod As Integer
    sDestinoDescripcion As String
    sNroCredito As String
    nSubDestinoCod As Integer
    sSubDestinoDescripcion As String
    nMontoP As Double
    nMontoS As Double
    nBeneficia As Integer
    nSumaCreditoVerde As Double
End Type
'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM Crédito Verde***
Public Type TEscalonamientoConf
    nMoraPorcentaje As Double
    nMontoCuoMayorA As Double
    nMontoCuoMenorIgual As Double
    nMontoCreMayorA As Double
    nMontoCreMenorIgual As Double
End Type
'END EJVG *******
'RECO20160706 ERS002-2016***************
Public Enum EstadoAutoExonera
    gEstadoPendiente = 0
    gEstadoAutorizado = 1
    gEstadoRechazado = 2
End Enum
Public Enum TipoNivelReg
    gTipoNivelAprob = 1
    gTipoNivelAuto = 2
End Enum
'RECO FIN *******************************
'RECO20161017 ERS060-2016****************
Public Enum TipoRegControl
    gTpoRegCtrlSolicitud = 1
    gTpoRegCtrlGarantia = 2
    gTpoRegCtrlInformeLegal = 3
    gTpoRegCtrlEvaluacion = 4
    gTpoRegCtrlSugerencia = 5
    gTpoRegCtrlRiesgos = 6
    gTpoRegCtrlMinutaLegal = 7
    gTpoRegCtrlNivAprobacion = 8
    gTpoRegCtrlAprobacion = 9
    gTpoRegCtrlAdmCreditos = 10
    gTpoRegCtrlDesembolso = 11
    gTpoRegCtrlAsistAgencia = 12 'ARLO20171002
End Enum
'RECO FIN *******************************
'*****-> LUCV2016915, Matrices Type para el Modulo de frmPersEstadosFinancierosDetalle <-*****
'MatrizDetalle->***********Activos/Pasivos/Flujo de Efectivo
Public Type tfrmPersEstadosFinancierosDetalle
    cDescripcion As String
    nImporte As Currency
End Type
'MatrizPrincipal->*********Activos/Pasivos/Flujo de Efectivo
Public Type tfrmPersEstadosFinancieros 'ACTIVOS
    cConcepto As String
    nImporte As Currency
    vTotal() As tfrmPersEstadosFinancierosDetalle 'Detalle en Matriz
    nConsCod As Integer
    nConsValor As Integer
    nConsValorGrupo As Integer
    nEstFinanTipo As Integer
End Type
'<-***** LUCV20160705

'WIOR 20180718 ***
Public Enum CodigoParamExtraExperian
      gCodFormulario = 1 'Código del formulario'
      gNumRepresentante = 2 'Número de Representantes Legales'
      gModalidad = 3 'Modalidad'
      gTipoCredito = 4 'Tipo de crédito'
      gTipoConsulta = 5 'Tipo de consulta'
      gTipoDocRepLegal1 = 6 'Tipo documento Representantes Lega 1'
      gDocRepLegal1 = 7 'Documento Representantes Lega 1'
      gTipoDocRepLegal2 = 8 'Tipo documento Representantes Lega 2'
      gDocRepLegal2 = 9 'Documento Representantes Lega 2'
      gTipoDocRepLegal3 = 10 'Tipo documento Representantes Lega 3'
      gDocRepLegal3 = 11 'Documento Representantes Lega 3'
      gTipoDocRepLegal4 = 12 'Tipo documento Representantes Lega 4'
      gDocRepLegal4 = 13 'Documento Representantes Lega 4'
      gEstCivilTitular = 14 'Estado civil Titular'
      gTipoDocConyuge = 15 'Tipo documento cónyuge titular'
      gDocConyuge = 16 'Documento cónyuge titular'
End Enum
'RECO FIN *********

'JOEP20190125 CP
Public Enum CodigoModulo
      gModSolicitud = 1 'Solicitud
      gModSugerencia = 2 'Sugerencia
      gModAmpliacion = 3 'Ampliacion
      gModRefinanciado = 4 'Refinanciado
      gModAprobacion = 5 'Aprobacion
      gModTodos = 6 'Todos
End Enum
'JOEP20190125 CP
'CTI5 ERS0012020**********************************
Public Enum TpoBienContrato
    gTpoBienConstruido = 1
    gTpoBienFuturo = 2
End Enum
'*************************************************
Public Enum OperacionesDineroElectronico
    gOpeTransferenciaCargo = 980003
    gOpeTransferenciaDeposito = 980004
End Enum

