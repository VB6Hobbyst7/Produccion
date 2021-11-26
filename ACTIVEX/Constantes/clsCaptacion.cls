VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptacion"
Attribute VB_GlobalNameSpace = False
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
End Enum

Public Enum CaptacTarjetaEstado
    gCapTarjEstActiva = 0   'A,N,K   Estado Tarjeta MagnetiCa Activa
    gCapTarjEstBloqueada = 1    'B   Estado Tarjeta Magnetica Bloqueada
    gCapTarjEstCancelada = 2    'C,V   Estado Tarjeta Magnetica Cancelada
    gCapTarjEstCmbioClave = 3
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

Public Enum CaptacTipoBloqueo
    gCapTpoBlqTotal = 11
    gCapTpoBlqRetiro = 12
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
End Enum

Public Enum CaptacOperacion
    'AHORROS
    gAhoApeEfec = 200101    'Apertura Efectivo
    gAhoApeChq = 200102     'Apertura Cheque
    gAhoApeTransf = 200103      'Apertura Transferencia
    gAhoApeLoteEfec = 200104
    gAhoApeLoteChq = 200105
    gAhoApeSobRemate = 200106
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
    
    gAhoCancAct = 200401 'Cancelación Activa
    gAhoCancInact = 200402 'Cancelación Inactiva
    gAhoCancTransfAct = 200403 'Cancelación Activa
    gAhoCancTransfInact = 200404 'Cancelación Inactiva
    gAhoCancInmovilizada = 200405 'Cancelacion Inmovilizada
    gAhoCancInmovFSD = 200406 'Cancelación Imnmovilizada pase FSD
    gAhoCancSobRemate = 200407  'Cancelacion por Sobrante de Remate
    
    gAhoCancSobregiroOP = 200408  'Cancelacion por Sobregiro de ordenes de Pago
    
    
    gAhoCapIntAct = 200501  'Capitalización Intereses Activa
    gAhoCapIntInact = 200502  'Capitalización Intereses Inactiva
    gAhoDctoEmiExt = 200601 'Descuento Emisión Extracto
    gAhoDctoEmiOP = 200602  'Descuento Emisión Chequera
    gAhoDctoInactiva = 200603  'Descuento Cuenta Inactiva
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
    gAhoTransAboCargIndebido = 200906  'Transferencia por Abono/Cargo Indebido
    gAhoTransAboIndebMNMN = 200907  'Transferencia MN - MN
    gAhoTransAboIndebMEME = 200908  'Transferencia ME - ME
    gAhoTransCargoIndeb = 200910  'Transferencia Cargo Indebido
    
    gAhoConsSaldo = 201001  'Consulta Saldo
    gAhoConsMovimiento = 201002 'Consulta Movimientos
    
    'PLAZO FIJO
    gPFApeEfec = 210101  'Apertura Efectivo
    gPFApeChq = 210102  'Apertura Cheque
    gPFApeTransf = 210103  'Apertura Transferencia
    gPFApeLoteEfec = 210104
    gPFApeLoteChq = 210105
    gPFRetInt = 210201  'Retiro de Intereses Efectivo
    gPFRetIntAboAho = 210202  'Retiro de Intereses Abono Cta Ahorros
    gPFCancEfec = 210301  'Cancelación
    gPFCancTransf = 210302  'Cancelación Transferencia
    gPFCapInt = 210401  'Capitalización Intereses
    gPFRenovacion = 210501  'Renovación
    gPFDupCert = 210601 'Duplicado de Certificado de Plazo Fijo
    gPFBusqCredOend = 210602 'Búsqueda de Crédito Pendiente
    gPFConsSaldo = 210701 'Consulta de Saldos
    gPFConsMovimiento = 210702 'Ultimos Movimientos
    
    'CTS
    gCTSApeEfec = 220101  'Apertura Efectivo
    gCTSApeChq = 220102  'Apertura Cheque
    gCTSApeTransf = 220103  'Apertura Transferencia
    gCTSApeLoteEfec = 220104
    gCTSApeLoteChq = 220105
    gCTSDepEfec = 220201    'Depósito Efectivo
    gCTSDepChq = 220202     'Depósito Cheque
    gCTSDepTransf = 220203      'Depósito Nota Abono
    gCTSDepPlanRRHH = 220204      'Depósito Planilla RRHH
    gCTSRetEfec = 220301      'Retiro Efectivo
    gCTSRetTransf = 220302  'Retiro Nota Cargo
    gCTSCancEfec = 220401      'Cancelación Efectivo
    gCTSCancTransf = 220402      'Cancelación Transferencia
    gCTSCapInt = 220501  'Capitalización Intereses
    gCTSConsSaldo = 220601 'Consulta de Saldos
    gCTSConsMovimiento = 220602 'Ultimos Movimientos
    
    gAhoExt = 230000  'EXTORNOS AHORROS
    gAhoExtApeEfec = 230101  'Extorno Apertura Efectivo
    gAhoExtApeChq = 230102  'Extorno Apertura Cheque
    gAhoExtApeTransf = 230103  'Extorno Apertura Transferencia
    gAhoExtApeLoteEfec = 230104  'Extorno Apertura Lote Cheque
    gAhoExtApeLoteChq = 230105  'Extorno Apertura Lote Efectivo
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
    
    
    gAhoExtCancAct = 230401  'Extorno Cancelación Activa
    gAhoExtCancTransfAct = 230402  'Extorno Cancelación Transferencia
    gAhoExtCancSobRemate = 230407  'Extorno Cancelacion Sobrante Remate
    gAhoExtDctoEmiExt = 230601 'Extorno Descuento Emisión Extracto
    gAhoExtDctoEmiOP = 230602 'Extorno Descuento Emisión Chequera
    gAhoExtTransferencia = 230900 'Extorno Transferencia entre cuentas
    gAhoExtTransAbono = 230901
    gAhoExtTransCargo = 230902
    
    gPFExt = 240000  'EXTORNOS PLAZO FIJO
    gPFExtApeEfec = 240101  'Extorno Apertura Efectivo
    gPFExtApeChq = 240102  'Extorno Apertura Cheque
    gPFExtApeTransf = 240103  'Extorno Apertura Transferencia
    gPFExtApeLoteEfec = 240104  'Extorno Apertura Lote Efectivo
    gPFExtApeLoteChq = 240105  'Extorno Apertura Lote Cheque
    gPFExtRetInt = 240201  'Extorno Retiro Intereses Efectivo
    gPFExtRetIntAboAho = 240202  'Extorno Retiro Intereses Abono Cta Ahorro
    gPFExtCancEfec = 240301  'Extorno Cancelación Efectivo
    gPFExtCancTransf = 240302  'Extorno Cancelación Transferencia
    
    gCTSExt = 250000  'EXTORNOS CTS
    gCTSExtApeEfec = 250101  'Extorno Apertura Efectivo
    gCTSExtApeChq = 250102  'Extorno Apertura Cheque
    gCTSExtApeTransf = 250103  'Extorno Apertura Transferencia
    gCTSExtApeLoteEfec = 250104  'Extorno Apertura Lote Efectivo
    gCTSExtApeLoteChq = 250105  'Extorno Apertura Lote Cheque
    gCTSExtDepEfec = 250201    'Extorno Depósito Efectivo
    gCTSExtDepChq = 250202     'Extorno Depósito Cheque
    gCTSExtDepTransf = 250203      'Extorno Depósito Transferencia
    gCTSExtRetEfec = 250301      'Extorno Retiro Efectivo
    gCTSExtRetTransf = 250302  'Extorno Retiro Nota Cargo
    gCTSExtCancEfec = 250401      'Extorno Cancelación Efectivo
    gCTSExtCancTransf = 250402      'Extorno Cancelación Transferencia
    
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
    
    
    'Promotores de Captaciones
    gCapRepVariosCapPromAho = 280709
    gCapRepVariosCapPromPF = 280710
    gCapRepVariosCapPromCTS = 280711
    
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
    
    'Otras Operaciones
    gOtrOpePagoFaltante = 300200
    gOtrOpeDepCtaBcoEfec = 300300
    
    'Ingresos
    gOtrOpeAhoOtrosIngresos = 300401
    gOtrOpeTransBancos = 300402
    gOtrOpeIngresosoCajaGeneral = 300403
    gOtrOpeComisionCMACLlam = 300404
    gOtrOpeTransferenciasEfect = 300405
    gOtrOpeRedondeoITF = 300499
    
    'Egresos
    gOtrOpeAhoOtrosEgresos = 300501
    gOtrOpeDepositoBancos = 300502
    gOtrOpeDevolucionCredPersonal = 300503
    
    'Tarifas
    gOtrOpeDuplicadoTarjeta = 300601
    gOtrOpeVentaListados = 300602
    gOtrOpeConsatnciaCancelacionCredito = 300603
    gOtrOpeElaboracionContrato = 300604
    
    gOtrOpe
    'Operaciones con Cheques
    gChqOpeRegistro = 900031
    gChqOpeModFechaValor = 900032
    gChqOpeValorInmediata = 900033
    gChqOpeConsultaEstado = 900034
    
    'Extornos Operaciones con Cheques
    gChqOpeExtRegistro = 900037
    gChqOpeExtModFechaValor = 900038
    gChqOpeExtValorInmediata = 900039
    
    'Extorno de OtrasOperaciones
    gOtrOpeExtorno = 390000
    
    'Giros
    gServGiroApertEfec = 310101
    gServGiroApertCargoCta = 310102
    gServGiroComiEfec = 310201
    gServGiroComiCargoCta = 310202
    gServGiroComiEfecOA = 310203
    gServGiroCancEfec = 310301
    gServGiroCancCargoCta = 310302
    
    
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
        
    gITFCobroCargo = 990101
    gITFCobroEfectivo = 990102
    gITFCobroEfectivoCred = 990103
    gITFCobroEfectivoPrend = 990104

    gITFCobroCMACCargo = 990301
    gITFCobroCMAC = 990302
    gITFCobroCMACCred = 990303
    gITFCobroCMACPrend = 990304

    gITFCobroCargoExt = 990201
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
    
    gConcITFCliente = 20
    gConcITFAsumido = 21
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
