VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColocacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    gColocConceptoCodGastoBusquedaRegistral = 1216
    gColocConceptoCodGastoComisionFlat = 1217
    gColocConceptoCodGastoProtesto1 = 1219
    gColocConceptoCodGastoProtesto2 = 1220
    gColocConceptoCodGastoGastosCobranza = 1221
    gColocConceptoCodGastoTasacion = 1222
    gColocConceptoCodGastoPenalidadPagoAdelantadoSoles = 1224
    gColocConceptoCodGastoPenalidadPagoAdelantadoDolares = 1225
    gColocConceptoCodGastoDsctoPlanillaSoles = 1227
    gColocConceptoCodGastoDsctoPlanillaDolares = 1228
    gColocConceptoCodGastoSuspenso = 1233
    gColocConceptoCodGastoVarios = 1299  'Gastos Varios
    
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
    gColPConceptoCodTasacion = 2200             'Pignoraticio: Tasacion
    gColPConceptoCodCustodia = 2201             'Pignoraticio: Custodia
    gColPConceptoCodCustodiaVencida = 2202      'Pignoraticio: Custodia Vencida
    gColPConceptoCodCustodiaDiferida = 2203     'Pignoraticio: Custodia Diferia
    gColPConceptoCodImpuesto = 2204             'Pignoraticio: Impuesto
    gColPConceptoCodPreparaRemate = 2205        'Pignoraticio: PreparaRemate
    gColPConceptoCodComisionRemate = 2206       'Pignoraticio: Comision Remate
    gColPConceptoCodSobranteRemate = 2207       'Pignoraticio: Sobrante Remate
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
End Enum

Enum ColocTiposGracia
    gColocTiposGraciaPrimeraCuota = 0
    gColocTiposGraciaUltimaCuota = 1
    gColocTiposGraciaProrateada = 2
    gColocTiposGraciaConfigurable = 3
    gColocTiposGraciaExonerada = 4
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
    gColPOpeCancelacEFE = 121200            'Col Pig : Cancelacion
        gColPOpeCancelNorEFE = 121201               'Col Pig : Cancelacion Nor EFEC
        gColPOpeCancelMorEFE = 121202               'Col Pig : Cancelacion Mor EFEC
    gColPOpeCancelacCHQ = 121210
        gColPOpeCancelNorCHQ = 121211               'Col Pig : Cancelacion Nor CHEQUE
        gColPOpeCancelMorCHQ = 121212               'Col Pig : Cancelacion Mor CHEQUE
    gColPOpeImpDuplicado = 121600        'Col Pig : Imprimir Duplicado
    gColPOpeDevJoyas = 121800            'Col Pig : Devol Joyas
    gColPOpeCobCusDiferida = 121900      'Col Pig : Cobrar Custod Diferida
    gColPOpeVtaRemate = 122000           'Col Pig : Venta Joyas en Remate
    gColPOpeAboSobCta = 122100           'Col Pig : Abono Sobrante Remate
    gColPOpePagSobrante = 122200         'Col Pig : Pago de Sobrante Remate
    gColPOpeAdjudica = 122500            'Col Pig : Adjudica Joyas
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
    gColPOpeExtRenov = 129100               'Col Pig : Extorno Renovacion
    gColPOpeExtDevJoyas = 129200            'Col Pig : Extorno Devol Joyas
    gColPOpeExtCustodDifer = 129300         'Col Pig : Extorno Custodia Difer
    gColPOpeExtCance = 129400               'Col Pig : Extorno Cancelacion
    gColPOpeExtDesemb = 129500              'Col Pig : Extorno Desembolso
    gColPOpeExtDupli = 129600               'Col Pig : Extorno Duplicado

    gColPOpeExtRenovEOCmac = 129801              'Col Pig : Extorno Renov En ot Cmac
    gColPOpeExtCanceEOCmac = 129802              'Col Pig : Extorno Cance En Ot Cmac

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
    gColCredRepCorrespClientes = 108390 'Correspondencia a Clientes
    gColCredRepCartaCobMoro1 = 108391 'Carta de Cobranza para Morosos 1
    gColCredRepCartaCobMoro2 = 108392 'Carta de Cobranza para Morosos 2
    gColCredRepCartaCobMoro3 = 108393 'Carta de Cobranza para Morosos 3
    gColCredRepCartaCobMoro4 = 108394 'Carta de Cobranza para Morosos 4
    gColCredRepCartaCobMoro5 = 108395 'Carta de Cobranza para Morosos 5
    gColCredRepCartaInvCredAlt = 108396 'Carta de Invitación Crédito Alterno
    gColCredRepCartaRecup = 108397 'Carta para Recuperaciones
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
        
    gColRecOpePagJudCD = 130300          'Recup : Pago Jud Con Demanda
        gColRecOpePagJudCDEfe = 130301        'Recup : Pago Jud CD Efectivo
        gColRecOpePagJudCDChq = 130302        'Recup : Pago Jud CD Cheque
        
    gColRecOpePagCast = 130400          'Recup : Pago Castigado
        gColRecOpePagCastEfe = 130401        'Recup : Pago Cast Efectivo
        gColRecOpePagCastChq = 130402        'Recup : Pago Cast Cheque
    
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
