Attribute VB_Name = "MColocPConstantes"
Option Explicit
'***************************************************************

Global Const geColPConceptoCodVtaNetaSubasta = 2211 'Pignoraticio: Venta Neta Subasta
Global Const geColPConceptoCodImpuestoVtaSubasta = 2212 'Pignoraticio: Impuesto Venta Subasta
Global Const geColPConceptoCodValorAdjudica = 2213 'Pignoraticio: Valor de Registro de Adjudicacion


Global Const geColPRecGarVentaSobranteCobrado = 4

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
Global Const gConsColPNroRematesParaAdjudic = 3045
Global Const gConsColPCostoNotifAvisoVenci = 3058 'DAOR 20070118


' Modificacion de Descripc Joyas
Global Const geColPModifDescJoyas = 120600

' Bloqueo de Joyas
Global Const geColPBloqueoJoyas = 120700

' Desbloqueo de Joyas
Global Const geColPDesbloqueoJoyas = 120800

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
'Global Const geColPCobCusDif = 121901
'Global Const geColPCobCusImp = 121902

' Venta Lotes en Remate
Global Const geColPVtaRemate = 122000

'Abono a cuenta de Ahorros
Global Const geColPAboSobCta = 122100

' Pago de Sobrantes
Global Const geColPPagSobrante = 122200
Global Const geColPPagSobraAdjudicado = 122300 ''*** PEAC 20090413
Global Const geColPPagSob = 122201

' Adjudicaciones de Creditos
Global Const geColPAdjudica = 122500

' Venta Adjudicados en Subasta
Global Const geColPVtaSubasta = 122800

'*********  OPERACIONES CON OTRAS CMACT
'**** Renovacion En Otra CMAC
Global Const geColPRenEnOtCj = 126100
'**** Renovacion En Otra CMAC - Morosa
Global Const geColPRenMorEnOtCj = 126200

'****  Cancelacion Contrato En Otra CMAC
Global Const geColPCanceEnOtCj = 126300

'**** Cancelacion Morosa En Otra CMAC
Global Const geColPCanMorEnOtCj = 126400

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

Public Function CalculaTasaEfectivaAnual(ByVal pnTasaMensual As Double) As Double
Dim lnTasa As Double
    lnTasa = (pnTasaMensual + 1) ^ 12 - 1
CalculaTasaEfectivaAnual = lnTasa
End Function

Public Function mfgEstadoCredPigDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColPEstRegis
            lsDesc = "Registrado"
        Case gColPEstDesem
            lsDesc = "Desembolsado"
        Case gColPEstDifer
            lsDesc = "Diferido (Para rescate)"
        Case gColPEstCance
            lsDesc = "Cancelado"
        Case gColPEstVenci
            lsDesc = "Vencido"
        Case gColPEstRemat
            lsDesc = "Remate"
        Case gColPEstPRema
            lsDesc = "Para Remate"
        Case gColPEstRenov
            lsDesc = "Renovado"
        Case gColPEstAdjud
            lsDesc = "Adjudicado"
        Case gColPEstSubas
            lsDesc = "Subastada"
        Case gColPEstAnula
            lsDesc = "Anulado"
        Case gColPEstChafa
            lsDesc = "Chafaloneado"
        Case gColocEstRech 'CROB20180602
            lsDesc = "Rechazado" 'CROB20180602
    End Select
    mfgEstadoCredPigDesc = lsDesc
End Function

