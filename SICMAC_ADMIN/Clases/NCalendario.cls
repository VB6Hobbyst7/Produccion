VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Tipos de Gracia
Enum TCalendTipoGracia
    PrimeraCuota = 1
    UltimaCuota = 2
    Prorateada = 3
    Configurable = 4
    Exonerada = 5
End Enum
'Tipos de Periodo
Enum TCalendTipoPeriodo
    PeriodoFijo = 1
    FechaFija = 2
End Enum

'Tipos de Cuota
Enum TCalendTipoCuota
    Fija = 1
    Creciente = 2
    Decreciente = 3
End Enum

Private Type TCalendario
    dFecha As Date
    NroCuota As Integer
    Cuota As Double
    Captital As Double
    IntComp As Double
    IntGra As Double
    Gasto As Double
    SaldoCap As Double
End Type

Dim Calendario() As TCalendario
Dim NroCuotas As Integer

Public Function GeneraGracia(ByVal pOptGracia As TCalendTipoGracia, ByVal pnInteres As Double, ByVal pnNrocuotas As Integer) As Variant
Dim nValorProrateado As Double
Dim I As Integer
Dim MatGracia() As Double
Dim sMatGracia() As String

    nValorProrateado = pnInteres / pnNrocuotas
    nValorProrateado = CDbl(Format(nValorProrateado, "#0.00"))
    ReDim MatGracia(pnNrocuotas)
    If pOptGracia = PrimeraCuota Then
        ReDim Preserve MatGracia(pnNrocuotas + 1)
        MatGracia(0) = pnInteres
    End If
    If pOptGracia = UltimaCuota Then
        ReDim Preserve MatGracia(pnNrocuotas + 1)
        MatGracia(pnNrocuotas) = pnInteres
    End If
    If pOptGracia = Prorateada Then
        For I = 0 To pnNrocuotas - 1
            If I = pnNrocuotas - 1 Then
                MatGracia(I) = pnInteres - CDbl(Format((nValorProrateado * (pnNrocuotas - 1)), "#0.00"))
            Else
                MatGracia(I) = nValorProrateado
            End If
        Next I
    End If
    
    If pOptGracia = Exonerada Then
        ReDim MatGracia(pnNrocuotas)
    End If
    
    ReDim sMatGracia(UBound(MatGracia))
    For I = 0 To UBound(MatGracia) - 1
        sMatGracia(I) = Format(MatGracia(I), "#0.00")
    Next I
    GeneraGracia = sMatGracia
End Function


Private Sub CalendarioCreciente(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNrocuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant)


Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim I As Integer
Dim oCredito As NCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nSumCuotas As Double

    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    nSumCuotas = 0
    
    For I = 1 To pnNrocuotas
        nSumCuotas = nSumCuotas + I
    Next I
    
    If pnTipoPeriodo = PeriodoFijo Then
        Set oCredito = New NCredito
        For I = 0 To pnNrocuotas - 1
            Calendario(I).dFecha = dDesembolso + pnPeriodo
            Calendario(I).NroCuota = I + 1
            Calendario(I).IntGra = 0#
            Calendario(I).Gasto = 0#
            Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            If I = pnNrocuotas - 1 Then
                Calendario(I).Captital = nSaldoCapital
            Else
                Calendario(I).Captital = CDbl(Format((Calendario(I).NroCuota / nSumCuotas) * pnMonto, "#0.00"))
            End If
            Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
            Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
            nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            dDesembolso = dDesembolso + pnPeriodo
        Next I
        
        
    Else
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New NCredito
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            nDia = pnDiaFijo
            For I = 0 To pnNrocuotas - 1
                If Not (I = 0 And pnDiaFijo > Day(dDesembolso) And Not bProxMes) Then
                    nMes = nMes + 1
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = 1
                    End If
                End If
                If nMes = 2 Then
                    If nDia > 28 Then
                        If nAnio Mod 4 = 0 Then
                            nDia = 29
                        Else
                            nDia = 28
                        End If
                    End If
                Else
                    If nDia > 30 Then
                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                            nDia = 30
                        End If
                    End If
                End If
                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                Calendario(I).dFecha = dFecTemp
                Calendario(I).NroCuota = I + 1
                Calendario(I).IntGra = 0#
                Calendario(I).Gasto = 0
                If I = 0 Then
                    Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(I).dFecha), nSaldoCapital)
                    Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(I).dFecha))
                Else
                    Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha), nSaldoCapital)
                    Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha))
                End If
                If I = pnNrocuotas - 1 Then
                    Calendario(I).Captital = nSaldoCapital
                Else
                    Calendario(I).Captital = CDbl(Format((Calendario(I).NroCuota / nSumCuotas) * pnMonto, "#0.00"))
                End If
                Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
                nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            Next I
        End If
    End If
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNrocuotas + 1)
            For I = pnNrocuotas To 1 Step -1
                Calendario(I) = Calendario(I - 1)
                Calendario(I).NroCuota = Calendario(I).NroCuota + 1
            Next I
            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            Calendario(0).NroCuota = 1
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Gasto = 0
            Calendario(0).IntComp = 0#
            Calendario(0).Cuota = Calendario(0).IntGra
            Calendario(0).Captital = 0#
            Calendario(0).SaldoCap = pnMonto
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For I = 0 To pnNrocuotas - 1
                Calendario(I).IntGra = MatGracia(I)
                Calendario(I).Cuota = Calendario(I).Cuota + MatGracia(I)
            Next I
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNrocuotas + 1)
            Calendario(pnNrocuotas).dFecha = Calendario(pnNrocuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNrocuotas).NroCuota = pnNrocuotas + 1
            Calendario(pnNrocuotas).IntGra = MatGracia(pnNrocuotas)
            Calendario(pnNrocuotas).Gasto = 0
            Calendario(pnNrocuotas).IntComp = 0#
            Calendario(pnNrocuotas).Cuota = Calendario(pnNrocuotas).IntGra
            Calendario(pnNrocuotas).Captital = 0#
            Calendario(pnNrocuotas).SaldoCap = 0#
        End If
    End If

End Sub

Private Sub CalendarioDecreciente(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNrocuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant)

Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim I As Integer
Dim oCredito As NCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer

    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    
            
    If pnTipoPeriodo = PeriodoFijo Then
        Set oCredito = New NCredito
        For I = 0 To pnNrocuotas - 1
            Calendario(I).dFecha = dDesembolso + pnPeriodo
            Calendario(I).NroCuota = I + 1
            Calendario(I).IntGra = 0#
            Calendario(I).Gasto = 0#
            Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            
            If I = pnNrocuotas - 1 Then
                Calendario(I).Captital = nSaldoCapital
            Else
                Calendario(I).Captital = CDbl(Format((pnMonto / pnNrocuotas), "#0.00"))
            End If
            Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
            Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
            nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            dDesembolso = dDesembolso + pnPeriodo
        Next I
    Else
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New NCredito
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            nDia = pnDiaFijo
            For I = 0 To pnNrocuotas - 1
                If Not (I = 0 And pnDiaFijo > Day(dDesembolso) And Not bProxMes) Then
                    nMes = nMes + 1
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = 1
                    End If
                End If
                If nMes = 2 Then
                    If nDia > 28 Then
                        If nAnio Mod 4 = 0 Then
                            nDia = 29
                        Else
                            nDia = 28
                        End If
                    End If
                Else
                    If nDia > 30 Then
                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                            nDia = 30
                        End If
                    End If
                End If
                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                Calendario(I).dFecha = dFecTemp
                Calendario(I).NroCuota = I + 1
                Calendario(I).IntGra = 0#
                Calendario(I).Gasto = 0
                If I = 0 Then
                    Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(I).dFecha), nSaldoCapital)
                    Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(I).dFecha))
                Else
                    Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha), nSaldoCapital)
                    Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha))
                End If
                If I = pnNrocuotas - 1 Then
                    Calendario(I).Captital = nSaldoCapital
                Else
                    Calendario(I).Captital = CDbl(Format((pnMonto / pnNrocuotas), "#0.00"))
                End If
                Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
                nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            Next I
        End If
    End If
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNrocuotas + 1)
            For I = pnNrocuotas To 1 Step -1
                Calendario(I) = Calendario(I - 1)
                Calendario(I).NroCuota = Calendario(I).NroCuota + 1
            Next I
            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            Calendario(0).NroCuota = 1
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Gasto = 0
            Calendario(0).IntComp = 0#
            Calendario(0).Cuota = Calendario(0).IntGra
            Calendario(0).Captital = 0#
            Calendario(0).SaldoCap = pnMonto
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For I = 0 To pnNrocuotas - 1
                Calendario(I).IntGra = MatGracia(I)
                Calendario(I).Cuota = Calendario(I).Cuota + MatGracia(I)
            Next I
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNrocuotas + 1)
            Calendario(pnNrocuotas).dFecha = Calendario(pnNrocuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNrocuotas).NroCuota = pnNrocuotas + 1
            Calendario(pnNrocuotas).IntGra = MatGracia(pnNrocuotas)
            Calendario(pnNrocuotas).Gasto = 0
            Calendario(pnNrocuotas).IntComp = 0#
            Calendario(pnNrocuotas).Cuota = Calendario(pnNrocuotas).IntGra
            Calendario(pnNrocuotas).Captital = 0#
            Calendario(pnNrocuotas).SaldoCap = 0#
        End If
    End If
End Sub

Private Sub CalendarioCuotaFija(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNrocuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant)
                
Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim I As Integer
Dim oCredito As NCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer

    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    If pnTipoPeriodo = PeriodoFijo Then
        Set oCredito = New NCredito
        For I = 0 To pnNrocuotas - 1
            Calendario(I).dFecha = dDesembolso + pnPeriodo
            Calendario(I).NroCuota = I + 1
            Calendario(I).IntGra = 0#
            Calendario(I).Gasto = 0#
            Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, pnPeriodo)
            If I = pnNrocuotas - 1 Then
                Calendario(I).Captital = nSaldoCapital
                Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
            Else
                Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
            End If
            Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
            nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
            dDesembolso = dDesembolso + pnPeriodo
        Next I
        
        
    Else
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New NCredito
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            nDia = pnDiaFijo
            For I = 0 To pnNrocuotas - 1
                If Not (I = 0 And pnDiaFijo > Day(dDesembolso) And (Not bProxMes)) Then
                    nMes = nMes + 1
                    If nMes > 12 Then
                        nAnio = nAnio + 1
                        nMes = 1
                    End If
                Else
                    If nMes = 2 Then
                        If nDia >= 29 Then
                            If nAnio Mod 4 <> 0 Then
                                nMes = nMes + 1
                            End If
                        End If
                    Else
                        If nDia > 30 Then
                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                                nMes = nMes + 1
                            End If
                        End If
                    End If
                End If
                If nMes = 2 Then
                    If nDia > 28 Then
                        If nAnio Mod 4 = 0 Then
                            nDia = 29
                        Else
                            nDia = 28
                        End If
                    End If
                Else
                    If nDia > 30 Then
                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
                            nDia = 30
                        End If
                    End If
                End If
                dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                Calendario(I).dFecha = dFecTemp
                Calendario(I).NroCuota = I + 1
                Calendario(I).IntGra = 0#
                Calendario(I).Gasto = 0
                If I = 0 Then
                    Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(I).dFecha), nSaldoCapital)
                    Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(I).dFecha))
                Else
                    Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha), nSaldoCapital)
                    Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha))
                End If
                If I = pnNrocuotas - 1 Then
                    Calendario(I).Captital = nSaldoCapital
                    Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
                Else
                    Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
                End If
                Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
                nSaldoCapital = nSaldoCapital - Calendario(I).Captital
            Next I
        End If
    End If
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNrocuotas + 1)
            For I = pnNrocuotas To 1 Step -1
                Calendario(I) = Calendario(I - 1)
                Calendario(I).NroCuota = Calendario(I).NroCuota + 1
            Next I
            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
            Calendario(0).NroCuota = 1
            Calendario(0).IntGra = MatGracia(0)
            Calendario(0).Gasto = 0
            Calendario(0).IntComp = 0#
            Calendario(0).Cuota = Calendario(0).IntGra
            Calendario(0).Captital = 0#
            Calendario(0).SaldoCap = pnMonto
        End If
        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
            For I = 0 To pnNrocuotas - 1
                Calendario(I).IntGra = MatGracia(I)
                Calendario(I).Cuota = Calendario(I).Cuota + MatGracia(I)
            Next I
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNrocuotas + 1)
            Calendario(pnNrocuotas).dFecha = Calendario(pnNrocuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNrocuotas).NroCuota = pnNrocuotas + 1
            Calendario(pnNrocuotas).IntGra = MatGracia(pnNrocuotas)
            Calendario(pnNrocuotas).Gasto = 0
            Calendario(pnNrocuotas).IntComp = 0#
            Calendario(pnNrocuotas).Cuota = Calendario(pnNrocuotas).IntGra
            Calendario(pnNrocuotas).Captital = 0#
            Calendario(pnNrocuotas).SaldoCap = 0#
        End If
    End If
End Sub

Private Sub ProcesarCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNrocuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant)
                
                Select Case pnTipoCuota
                    Case Creciente
                        Call CalendarioCreciente(pnMonto, pnTasaInt, pnNrocuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia)
                    Case Decreciente
                        Call CalendarioDecreciente(pnMonto, pnTasaInt, pnNrocuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia)
                    Case Fija
                        Call CalendarioCuotaFija(pnMonto, pnTasaInt, pnNrocuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia)
                End Select
End Sub
                
Public Function GeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNrocuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant) As Variant
Dim sCalendPagos() As String
Dim I As Integer

        On Error GoTo ErrorGeneraCalendario
        ReDim Calendario(pnNrocuotas)
        Call ProcesarCalendario(pnMonto, pnTasaInt, pnNrocuotas, pnPeriodo, pdFecDesemb, pnTipoCuota, _
                pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia)
        ReDim sCalendPagos(UBound(Calendario), 8)
        For I = 0 To UBound(Calendario) - 1
            sCalendPagos(I, 0) = Format(Calendario(I).dFecha, "dd/mm/yyyy")
            sCalendPagos(I, 1) = Trim(Str(Calendario(I).NroCuota))
            sCalendPagos(I, 2) = Format(Calendario(I).Cuota, "#0.00")
            sCalendPagos(I, 3) = Format(Calendario(I).Captital, "#0.00")
            sCalendPagos(I, 4) = Format(Calendario(I).IntComp, "#0.00")
            sCalendPagos(I, 5) = Format(Calendario(I).IntGra, "#0.00")
            sCalendPagos(I, 6) = Format(Calendario(I).Gasto, "#0.00")
            sCalendPagos(I, 7) = Format(Calendario(I).SaldoCap, "#0.00")
        Next I
        GeneraCalendario = sCalendPagos
        Exit Function

ErrorGeneraCalendario:
        MsgBox Err.Description, vbCritical, "Aviso"
        
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
