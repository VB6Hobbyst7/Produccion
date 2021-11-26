VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColPCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Clase que realiza los calculos financieros asociados a
'** Credito Pignoraticio
Option Explicit
Dim oError As New COMConecta.COMErrorHandling

Public Function nCalculaValorTasacion(ByVal pnPOro14 As Double, ByVal pnPOro16 As Double, ByVal pnPOro18 As Double, ByVal pnPOro21 As Double) As Double
' Calcula el valor de tasación
' Calcula en base al precio del oro en el mercado

Dim loColPCalculos As COMDColocPig.DCOMColPCalculos
Dim lnPrecioOro14 As Double
Dim lnPrecioOro16 As Double
Dim lnPrecioOro18 As Double
Dim lnPrecioOro21 As Double
Dim lnValTasacion As Double

On Error GoTo ECalculaValorT

Set loColPCalculos = New COMDColocPig.DCOMColPCalculos
    lnPrecioOro14 = loColPCalculos.dObtieneColocParametro(gColPConceptoCodOro14)
    lnPrecioOro16 = loColPCalculos.dObtieneColocParametro(gColPConceptoCodOro16)
    lnPrecioOro18 = loColPCalculos.dObtieneColocParametro(gColPConceptoCodOro18)
    lnPrecioOro21 = loColPCalculos.dObtieneColocParametro(gColPConceptoCodOro21)
Set loColPCalculos = Nothing

'********** Realiza el calculo
lnValTasacion = (pnPOro14 * lnPrecioOro14) + (pnPOro16 * lnPrecioOro16) + (pnPOro18 * lnPrecioOro18) + (pnPOro21 * lnPrecioOro21)
'*********
nCalculaValorTasacion = lnValTasacion
Exit Function

ECalculaValorT:
    Call oError.RaiseError(Err.Number, "Calcula Valor Tasacion <<NColPCalculos>>")

End Function

Public Function nCalculaCostoTasacion(ByVal pnValorTasacion As Currency, ByVal pnTasaTasacion As Double) As Currency
' Calcula el Costo de tasación

On Error GoTo ECalculaCostoT
'********** Realiza el calculo
nCalculaCostoTasacion = pnValorTasacion * pnTasaTasacion

Exit Function

ECalculaCostoT:
    Call oError.RaiseError(Err.Number, "Calcula Costo Tasacion <<NColPCalculos>>")

End Function

Public Function nCalculaCostoCustodia(ByVal pnValorTasacion As Currency, ByVal pnTasaCustodia As Double, _
    ByVal pnPlazo As Integer) As Currency

On Error GoTo ECalculaCostoC

'********** Realiza el calculo
nCalculaCostoCustodia = pnValorTasacion * pnTasaCustodia * (pnPlazo / 30)

Exit Function

ECalculaCostoC:
    Call oError.RaiseError(Err.Number, "Calcula Costo Custodia <<NColPCalculos>>")

End Function

Public Function nCalculaCostoCustodiaMoratorio(ByVal pnValorTasacion As Currency, _
       ByVal pnTasaCustodiaMorat As Double, ByVal pnDiasAtraso As Integer) As Currency
       
On Error GoTo ECalculo

'********** Realiza el calculo
nCalculaCostoCustodiaMoratorio = pnValorTasacion * (((1 + pnTasaCustodiaMorat) ^ (pnDiasAtraso / 30)) - 1)

Exit Function

ECalculo:
    Call oError.RaiseError(Err.Number, "Calcula Costo Custodia Moratorio <<NColPCalculos>>")
End Function

Public Function nCalculaCostoCustodiaDiferida(ByVal pnValorTasacion As Currency, _
        ByVal pnDiasTranscurridos As Integer, ByVal pnPorcentajeCustodiaDiferida As Double, _
        ByVal pnIGV As Double) As Currency
On Error GoTo ECalculo

'********** Realiza el calculo
If pnDiasTranscurridos < 30 Then
    nCalculaCostoCustodiaDiferida = Format((pnPorcentajeCustodiaDiferida / 30) * (1 + pnIGV) * pnValorTasacion * (0), "#0.00")
Else
    nCalculaCostoCustodiaDiferida = Format((pnPorcentajeCustodiaDiferida / 30) * (1 + pnIGV) * pnValorTasacion * (pnDiasTranscurridos - 30), "#0.00")
End If

Exit Function

ECalculo:
    Call oError.RaiseError(Err.Number, "Calcula Costo Custodia Moratorio <<NColPCalculos>>")
End Function
Public Function nCalculaImpuestoDesembolso(ByVal pnCostoTasacion As Currency, ByVal pnInteres As Currency, _
    ByVal pnCostoCustodia As Currency, ByVal pnTasaImpuesto As Double) As Currency

On Error GoTo ECalculo
'********** Realiza el calculo

nCalculaImpuestoDesembolso = (pnCostoTasacion + pnInteres + pnCostoCustodia) * pnTasaImpuesto
'*********

Exit Function

ECalculo:
    Call oError.RaiseError(Err.Number, "Calcula Impuesto Desembolso <<NColPCalculos>>")

End Function

Public Function nCalculaInteresAdelantado(ByVal pnSaldoCapital As Currency, ByVal pnTasaInteres As Double, _
        ByVal pnPlazo As Integer) As Currency

On Error GoTo ECalculo
'********** Realiza el calculo
nCalculaInteresAdelantado = pnSaldoCapital * (1 - (1 / ((1 + pnTasaInteres / 100) ^ (pnPlazo / 30))))
'*********

Exit Function

ECalculo:
    Call oError.RaiseError(Err.Number, "Calculo Interes Adelantado <<NColPCalculos>>")
End Function

Public Function nCalculaInteresMoratorio(ByVal pnSaldoCapital As Currency, ByVal pnTasaInteresMorat As Double, _
        ByVal pnDiasAtraso As Integer, Optional ByVal nIntCompen As Currency = 0) As Currency

On Error GoTo ECalculo
'pnTasaInteresMorat es mensual, convertida a Anual
'Dim lnTasaAnual As Double
'lnTasaAnual = (pnTasaInteresMorat / 100 + 1) ^ 12 - 1


'********** Realiza el calculo
'nCalculaInteresMoratorio = pnSaldoCapital * ((1 + lnTasaAnual) ^ (pnDiasAtraso / 360) - 1)
'*********
'********* Formula Interes Moratorio CMAC Santa

'*** PEAC 20170517
'nCalculaInteresMoratorio = pnSaldoCapital * ((pnTasaInteresMorat / 100 + 1) ^ (pnDiasAtraso / 30) - 1)
nCalculaInteresMoratorio = (pnSaldoCapital + nIntCompen) * ((pnTasaInteresMorat / 100 + 1) ^ (pnDiasAtraso / 30) - 1)
'*********

Exit Function

ECalculo:
    Call oError.RaiseError(Err.Number, "Calculo Interes Moratorio <<NColPCalculos>>")
End Function

Public Function nCalculaDeudaPignoraticio(ByVal pnSaldoCapital As Currency, ByVal pdFecVenc As Date, _
    ByVal pnValorTasac As Currency, ByVal pnTasaIntVenc As Double, ByVal pnTasaCustodiaVenc As Double, _
    ByVal pnTasaImpuesto As Double, ByVal pnEstado As ColocEstado, ByVal pnTasaPreparaRemate As Double, _
    ByVal pdFecCalculoDeuda As Date) As Currency

Dim lnDiasAtra As Integer
Dim lnDeuda As Currency
Dim lnIntMorat As Currency
Dim lnImpuesto As Currency
Dim lnCostoCustodiaMorat As Currency
'If Len(ldFecParaDeuda) <= 0 Then ldFecParaDeuda = gdFecSis

lnDiasAtra = IIf(DateDiff("d", pdFecVenc, pdFecCalculoDeuda) <= 0, 0, DateDiff("d", pdFecVenc, pdFecCalculoDeuda))
If lnDiasAtra = 0 Then
   lnIntMorat = 0
   lnCostoCustodiaMorat = 0
   lnImpuesto = 0
Else
   lnIntMorat = nCalculaInteresMoratorio(pnSaldoCapital, pnTasaIntVenc, lnDiasAtra)
   lnIntMorat = Round(lnIntMorat, 2)
   '*
   lnCostoCustodiaMorat = nCalculaCostoCustodiaMoratorio(pnValorTasac, pnTasaCustodiaVenc, lnDiasAtra)
   lnCostoCustodiaMorat = Round(lnCostoCustodiaMorat, 2)
   '*
   lnImpuesto = (lnIntMorat + lnCostoCustodiaMorat) * pnTasaImpuesto
   lnImpuesto = Round(lnImpuesto, 2)
End If
lnDeuda = Round(pnSaldoCapital, 2) + lnIntMorat + lnCostoCustodiaMorat + lnImpuesto

If pnEstado = gColPEstPRema Or pnEstado = gColPEstAdjud Then     ' Si esta en via de Remate
   lnDeuda = lnDeuda + Round((pnTasaPreparaRemate * pnValorTasac), 2)
End If
nCalculaDeudaPignoraticio = lnDeuda

End Function

Public Function nCalculaMinimoPagar(ByVal pnSaldoCap As Currency, ByVal pnTasaInteres As Double, _
        ByVal pnPlazo As Integer, ByVal pnTasaCustodia As Double, ByVal pdFecVenc As Date, _
        ByVal pnValorTasac As Currency, ByVal pnTasaIntVenc As Double, ByVal pnTasaCustodiaVenc As Double, _
        ByVal pnTasaImpuesto As Double, ByVal pnEstado As ColocEstado, ByVal pnTasaPreparaRemate As Double, _
        ByVal pdFecCalculo As Date) As Currency

' ** Calculo del Monto Minimo a Pagar **
Dim lnDiasAtra As Integer
Dim lnIntMorat As Currency
Dim lnCostoCustodiaMorat As Currency
Dim lnCostoPreparaRemate As Currency
Dim lnFactor As Currency
Dim lnInteres As Currency
Dim lnCostoCustodia As Currency
Dim lnImpuesto As Currency
Dim lnMinimo As Currency

lnDiasAtra = IIf(DateDiff("d", pdFecVenc, pdFecCalculo) <= 0, 0, DateDiff("d", pdFecVenc, pdFecCalculo))
If lnDiasAtra = 0 Then
   lnIntMorat = 0
   lnCostoCustodiaMorat = 0
Else
   lnIntMorat = nCalculaInteresMoratorio(pnSaldoCap, pnTasaIntVenc, lnDiasAtra)
   lnIntMorat = Round(lnIntMorat, 2)
   '*
   lnCostoCustodiaMorat = nCalculaCostoCustodiaMoratorio(pnValorTasac, pnTasaCustodiaVenc, lnDiasAtra)
   lnCostoCustodiaMorat = Round(lnCostoCustodiaMorat, 2)
End If
If pnEstado = gColPEstPRema Or pnEstado = gColPEstAdjud Then     ' Si esta en via de Remate
    lnCostoPreparaRemate = Round((pnTasaPreparaRemate * pnValorTasac), 2)
Else
    lnCostoPreparaRemate = 0
End If

lnFactor = nCalculaFactorRenovacion(pnTasaInteres, pnPlazo)
lnInteres = pnSaldoCap * lnFactor
lnInteres = Round(lnInteres, 2)
lnCostoCustodia = nCalculaCostoCustodia(pnValorTasac, pnTasaCustodia, pnPlazo)
lnCostoCustodia = Round(lnCostoCustodia, 2)

lnImpuesto = (lnIntMorat + lnCostoCustodiaMorat + lnInteres + lnCostoCustodia) * pnTasaImpuesto
lnImpuesto = Round(lnImpuesto, 2)

lnMinimo = lnIntMorat + lnCostoCustodiaMorat + lnInteres + lnCostoCustodia + lnImpuesto + lnCostoPreparaRemate
nCalculaMinimoPagar = Round(lnMinimo, 2)
nCalculaMinimoPagar = lnMinimo

End Function

Public Function nCalculaFactorRenovacion(ByVal pnTasaInteres As Double, ByVal pnPlazo As Integer) As Double
On Error GoTo ECalculo
'pnTasaInteres es mensual, convertida a Anual
Dim lnTasaAnual As Double
lnTasaAnual = (pnTasaInteres / 100 + 1) ^ 12 - 1
'********** Realiza el calculo
nCalculaFactorRenovacion = 1 - 1 / (1 + lnTasaAnual) ^ (pnPlazo / 360)
'*********

Exit Function

ECalculo:
    Call oError.RaiseError(Err.Number, "Calcula Factor Renovacion <<NColPCalculos>>")

End Function

Public Function nCalculaPrecioVentaRemate(ByVal pnDeuda As Currency, ByVal pnPrecioBase As Currency, _
                ByVal pnFactorPrecioVenta As Double) As Currency
Dim lnValor As Currency

If pnDeuda > pnPrecioBase Then
    lnValor = Round(pnDeuda, 0)
Else
    lnValor = Round(pnPrecioBase, 0)
End If

'If gsCodCMAC = "112" Then  'Trujillo
'    If (lnValor - Int(lnValor)) = 0 Then  ' Si es entero
'        If (lnValor Mod 5) = 0 Then
'            nCalculaPrecioVentaRemate = lnValor
'        Else
'            nCalculaPrecioVentaRemate = (Int(lnValor / 5) * 5) + 5
'        End If
'    Else    'Si no es entero
'        nCalculaPrecioVentaRemate = (Int(lnValor / 5) * 5) + 5
'    End If
'Else  ' Santa
    nCalculaPrecioVentaRemate = lnValor
'End If
End Function

Public Function nCalculaValorAdjudicacion(ByVal pnSaldoCap As Currency, ByVal pnValorMercado As Currency) As Currency
Dim lnValor As Currency
' Adjudica al menor entre: Saldo Capital - Valor de Mercado
lnValor = IIf(pnSaldoCap > pnValorMercado, pnValorMercado, pnSaldoCap)
lnValor = Round(lnValor, 2)
nCalculaValorAdjudicacion = lnValor
End Function

Public Function nCalculaPrecioVentaSubasta(ByVal pnPrecioBase As Currency, _
                ByVal pnTasaIGV As Double) As Currency
Dim lnValor As Currency

lnValor = pnPrecioBase * (1 + pnTasaIGV)
lnValor = Round(pnPrecioBase, 2)

If (lnValor - Int(lnValor)) = 0 Then  ' Si es entero
    If (lnValor Mod 5) = 0 Then
        nCalculaPrecioVentaSubasta = lnValor
    Else
        nCalculaPrecioVentaSubasta = (Int(lnValor / 5) * 5) + 5
    End If
Else    'Si no es entero
    nCalculaPrecioVentaSubasta = (Int(lnValor / 5) * 5) + 5
End If
End Function

Public Function nCalculaGastosCorrespondencia(ByVal psctacod As String, ByRef psmensaje As String) As Currency
Dim lsSQL As String
Dim loReg As COMConecta.DCOMConecta
Dim lrDatos As ADODB.Recordset

Dim lnValor As Currency
lnValor = 0
lsSQL = "SELECT Sum(nMonto - nMontoPagado) as GastoC FROM ColocCalendDetPig " & _
        " WHERE cCtaCod ='" & psctacod & "' And nNroCalen = 1 And nColocCalendApl = 0 " & _
        " And nCuota = 1 And nPrdConceptoCod = 2215 "

Set loReg = New COMConecta.DCOMConecta
loReg.AbreConexion
Set lrDatos = loReg.CargaRecordSet(lsSQL)
    
    If lrDatos Is Nothing Then
        psmensaje = "ERROR: al Buscar Gastos de Correspondencia"
        Exit Function
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        psmensaje = "ERROR: al Buscar Gastos de Correspondencia"
        Exit Function
    End If
     
    lnValor = Format(IIf(IsNull(lrDatos!GastoC), 0, lrDatos!GastoC), "#,#0.00")
    nCalculaGastosCorrespondencia = lnValor
    
Set lrDatos = Nothing
    
End Function

'*** PEAC 20080806
Public Function nCalculaInteresAlVencimiento(ByVal pnSaldoCapital As Currency, ByVal pnTasaInteres As Double, _
        ByVal pnPlazo As Integer) As Currency

On Error GoTo ECalculo
'********** Realiza el calculo
'Formula establecida en el Acta Nº 164-2008/TI-D
nCalculaInteresAlVencimiento = (((1 + pnTasaInteres / 100) ^ (pnPlazo / 30)) - 1) * pnSaldoCapital
'*********

Exit Function

ECalculo:
    Call oError.RaiseError(Err.Number, "Calculo Interes al Vencimiento <<NCOMColPCalculos>>")
End Function

