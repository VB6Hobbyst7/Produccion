Attribute VB_Name = "gColPConstantes"
Option Explicit

'************************************************************************************
'Colores para montos
'Fondo.
' Soles -> Blanco
Global Const pColFonSoles = &H80000005
' Dolares -> Verde
Global Const pColFonDolares = &HFF00&
'Primer Plano
' Ingreso -> Negro
Global Const pColPriIngreso = &HFF0000   '&H0&
' Egreso -> Rojo
Global Const pColPriEgreso = &HFF&
'************************************************************************************

'Constantes de Conceptos
Global Const geColPConceptoCodCapital = 2000  'Pignoraticio: Capital
Global Const geColPConceptoCodCapitalVencido = 2001  'Pignoraticio: Capital Vencido
Global Const geColPConceptoCodInteresCompensatorio = 2100 'Pignoraticio: Interes Compensatorio
Global Const geColPConceptoCodInteresMoratorio = 2101  'Pignoraticio: Interes Moratorio
Global Const geColPConceptoCodTasacion = 2200 'Pignoraticio: Tasacion
Global Const geColPConceptoCodCustodia = 2201  'Pignoraticio: Custodia
Global Const geColPConceptoCodCustodiaVencida = 2202  'Pignoraticio: Custodia Vencida
Global Const geColPConceptoCodCustodiaDiferida = 2203  'Pignoraticio: Custodia Diferia
Global Const geColPConceptoCodImpuesto = 2204 'Pignoraticio: Impuesto
Global Const geColPConceptoCodPreparaRemate = 2205 'Pignoraticio: PreparaRemate
Global Const geColPConceptoCodComisionRemate = 2206 'Pignoraticio: Comision Remate
Global Const geColPConceptoCodSobranteRemate = 2207 'Pignoraticio: Sobrante Remate
Global Const geColPConceptoCodCustodiaDiferidaImpuesto = 2208 ' Pignoraticio : Custodia Diferida Impuesto

Global Const geColPConceptoCodCostoDuplicado = 2210 'Pignoraticio: Impuesto

Global Const geColPConceptoCodVtaNetaSubasta = 2211 'Pignoraticio: Venta Neta Subasta
Global Const geColPConceptoCodImpuestoVtaSubasta = 2212 'Pignoraticio: Impuesto Venta Subasta
Global Const geColPConceptoCodValorAdjudica = 2213 'Pignoraticio: Valor de Registro de Adjudicacion

Global Const geColPConceptoCodOro14 = 2251 'Pignoraticio: Oro 14k
Global Const geColPConceptoCodOro16 = 2252 'Pignoraticio: Oro 16k
Global Const geColPConceptoCodOro18 = 2253 'Pignoraticio: Oro 18k
Global Const geColPConceptoCodOro21 = 2254 'Pignoraticio: Oro 21k

Global Const geColPConceptoCodMontoEntregar = 2271 'Pignoraticio: Monto Entregar

Global Const geColPConceptoCodCambioCarteraNomalMoroso = 2301 'Pignoraticio: Cambio Cartera Nomal-Moroso
Global Const geColPConceptoCodCambioCarteraMorosoNormal = 2302 'Pignoraticio: Cambio Cartera Moroso-Normal
Global Const geColPConceptoCodCambioTasacionNormalMoroso = 2311 'Pignoraticio: Cambio Tasacion Normal-Moroso
Global Const geColPConceptoCodCambioTasacionMorosoNormal = 2312 'Pignoraticio: Cambio Tasacion Moroso-Normal
'******
Global Const geColCJConceptoCodJudGasto01 = 1331  'Cobranza Judicial: Gasto 01
Global Const geColPConceptoCodJudGastoVarios = 1332  'Cobranza Judicial: Gastos Varios
Global Const geColPConceptoCodJudComision = 1341  'Cobranza Judicial: Comision 01
Global Const geColPConceptoCodJudComision99 = 1342 'Cobranza Judicial: Comision Varios

' Contantes de Estados de Proceso de Recuperación de Garantias
Global Const geColPRecGarEstNoIniciado = 0
Global Const geColPRecGarEstIniciado = 1
Global Const geColPRecGarEstTerminado = 2

' Contantes de Estados Creditos en el Proceso de Venta
Global Const geColPRecGarVentaEnProceso = 0
Global Const geColPRecGarVentaVendido = 1
Global Const geColPRecGarVentaNoVendido = 2
Global Const geColPRecGarVentaSobranteCobrado = 4
Global Const geColPRecGarAdjudicado = 5

'Constante de Parametros
Global Const gConsColPTasaCustodia = 3001
Global Const gConsColPTasaTasacion = 3002
Global Const gConsColPTasaImpuesto = 3003
Global Const gConsColPTasaCustodiaVencida = 3004
Global Const gConsColPTasaPreparaRemate = 3005
Global Const gConsColPTasaComisionRemate = 3006
Global Const gConsColPTasaIGV = 3007

Global Const gConsColPPorcentajePrestamo = 3011
Global Const gConsColPMaxCostoTasacion = 3012
Global Const gConsColPMinPesoOro = 3013
Global Const gConsColPLim1MontoPrestamo = 3014
Global Const gConsColPLim2MontoPrestamo = 3015
Global Const gConsColPToleranciaMontoPrestamo = 3016

Global Const gConsColPNroImpresionesContrato = 3017
Global Const gConsColPMaximoDiasDesembolso = 3018
Global Const gConsColPPrecioOro10 = 3021
Global Const gConsColPPrecioOro12 = 3022
Global Const gConsColPPrecioOro14 = 3023
Global Const gConsColPPrecioOro16 = 3024
Global Const gConsColPPrecioOro18 = 3025
Global Const gConsColPPrecioOro21 = 3026

Global Const gConsColPMaxNroRenovac = 3031
Global Const gConsColPCostoDuplicadoContrato = 3032
Global Const gConsColPMaxDiasCustodiaDiferida = 3033
Global Const gConsColPPorcentajeCustodiaDiferida = 3034
Global Const gConsColPFactorPrecioBaseRemate = 3041
Global Const gConsColPDiasAtrasoParaRemate = 3042
Global Const gConsColPDiasAtrasoCartaVenc = 3043
'*** PEAC 20080515
Global Const gConsColPDiasAtrasoParaAdjudicar = 3044

Global Const gConsColPNroRematesParaAdjudic = 3045

Global Const gConsColPDiasCambioCartera = 3050
Global Const gConsColPCostoNotifAvisoVenci = 3058 'DAOR

' Contantes de Procesos

' Registro Contrato Pignoraticio
Global Const geColPRegContrato = 120100 ' **


' Desembolso de Credito Pignoraticio
Global Const geColPDesembolso = 120200

' Anulacion de Credito Pignoraticio
Global Const geColPAnula = 120300

' Anulacion de Credito No Desembolsado
Global Const geColPAnulNoDesemb = 120400

' Devolucion de Joyas No Desembolsadas
Global Const geColPDevJoyasNoDesemb = 120500

' Modificacion de Descripc Joyas
Global Const geColPModifDescJoyas = 120600

' Bloqueo de Joyas
Global Const geColPBloqueoJoyas = 120700

' Desbloqueo de Joyas
Global Const geColPDesbloqueoJoyas = 120800

' Renovacion de Credito Pignoraticio
Global Const geColPRenovac = 121000


' Renovacion de Credito Pignoraticio - Morosa
Global Const geColPRenovMorosa = 121100


Global Const geColPRenMorCambCart = 121108

' Cancelacion de Credito Pignoraticio
Global Const geColPCancelac = 121200

' Cancelacion Moroso de Credito Pignoraticio
Global Const geColPCancelMorosa = 121300

' Imprimir Duplicado
Global Const geColPImpDuplicado = 121600

' Devolucion de Joyas
Global Const geColPDevJoyas = 121800

' Cobrar Custodia Diferida
Global Const geColPCobCusDiferida = 121900

' Venta Lotes en Remate
Global Const geColPVtaRemate = 122000

'Abono a cuenta de Ahorros
Global Const geColPAboSobCta = 122100

' Pago de Sobrantes
Global Const geColPPagSobrante = 122200
Global Const geColPPagSob = 122201

'PEAC 20090316 - PAGO SOBRANTE DE ADJUDICACION
Global Const geColPPagSobraAdjudica = 122300

' Adjudicaciones de Creditos
Global Const geColPAdjudica = 122500

' Venta Adjudicados en Subasta
Global Const geColPVtaSubasta = 122800

'*********  OPERACIONES CON OTRAS CMACT
'**** Renovacion En Otra CMAC
Global Const geColPRenEnOtCj = 126100

'****  Cancelacion Contrato En Otra CMAC
Global Const geColPCanceEnOtCj = 126200

'****  Cancelacion Contrato En Otra CMAC
Global Const geColPAmortEnOtCj = 126300

'**** Renovacion DE Otra CMAC
Global Const geColPRenDEOtCj = 127100

'****  Cancelacion DE Otra CMAC
Global Const geColPCanDEOtCj = 127200

' Cambio de Cartera en el Cierre Diario
Global Const geColPPigCamCarNorMor = 125400
Global Const geColPPigCanValTasVigVen = 125401

'****** CONSTANTES DE EXTORNOS
'****  Desembolso de contrato
Global Const geColPExtDesemb = 129500
'****  Duplicado de contrato
Global Const geColPExtDupli = 129600

'****  Cancelación/Rescate de contrato
Global Const geColPExtCance = 129000

'****  Renovación de contrato
Global Const geColPExtRenov = 129100

'****  Devolución de Prendas
Global Const geColPExtDevJoyas = 129200

'****  Custodia Diferida
Global Const geColPExtCustodDifer = 129300

'***** LLAMADA RENOV  DE OTRA CMAC
Global Const geColPExtRenovDeOtCj = 129801
'***** LLAMADA CANCEL  DE OTRA CMAC
Global Const geColPExtCancelDeOtCj = 129802

