VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPigCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CalcValorTasacion(ByVal pnPPEstJoya As Double, ByVal pnPesoN As Double, _
    ByVal pnPrecio As Currency) As Currency

Dim lnValorJoya As Currency
Dim lnValorTasac As Currency
    
    lnValorJoya = pnPesoN * pnPrecio
    lnValorTasac = (lnValorJoya * pnPPEstJoya / 100)
    CalcValorTasacion = lnValorTasac

End Function

Public Function CalcValorPrestamo(ByVal pnPPTipoClte As Double, ByVal pnValorTasacion As Currency) As Currency
Dim lnValorPrest As Currency
    
    lnValorPrest = pnPPTipoClte * pnValorTasacion / 100
    CalcValorPrestamo = lnValorPrest

End Function

Public Function CalcPagoMin(ByVal pnPorMin As Double, ByVal pnCapital As Currency, ByVal pnComSer As Double) As Currency
Dim lnPagoMin As Currency

    lnPagoMin = (pnCapital * pnPorMin / 100) + pnComSer
    CalcPagoMin = lnPagoMin
    
End Function

Public Function nCalculaIntCompensatorio(ByVal pnSaldoCapital As Currency, ByVal pnTasaInteres As Double, _
        ByVal pnPlazo As Long) As Currency

On Error GoTo ErrorCalculo

    nCalculaIntCompensatorio = (((1 + pnTasaInteres / 100) ^ (pnPlazo / 30)) - 1) * pnSaldoCapital
Exit Function

ErrorCalculo:
    Call RaiseError(Err.Number, "Calculo Interes Compensatorio <<NPigCalculos>>")
End Function

Public Function nCalculaConcepto(ByVal pnTpoValor As Integer, ByVal pnValor As Currency, ByVal pnMontoMin As Currency, _
                                                 ByVal pnMontoMax As Currency, ByVal pnSaldoCap As Currency)
Dim lnValor As Currency

On Error GoTo ErrorCalculo
           
     If pnTpoValor = 1 Then    'Monto
        lnValor = pnValor
     ElseIf pnTpoValor = 2 Then    'Porcentaje
        lnValor = Round(pnValor * pnSaldoCap / 100, 2)       'Si Cancela Antes del Plazo establecido
        If pnMontoMin > lnValor Then
           lnValor = pnMontoMin
        End If
        If pnMontoMax < lnValor Then
           lnValor = pnMontoMax
        End If
     End If

     nCalculaConcepto = lnValor

Exit Function

ErrorCalculo:
    Call RaiseError(Err.Number, "Calculo del Valor del Concepto <<NPigCalculos>>")

End Function

Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double

If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If

End Function

