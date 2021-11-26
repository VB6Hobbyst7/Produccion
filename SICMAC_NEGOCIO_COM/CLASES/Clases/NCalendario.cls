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

'Para Reporte no Borrar
Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String

Dim Calendario() As TCalendario
Dim NroCuotas As Integer

Public Function GeneraGracia(ByVal pOptGracia As TCalendTipoGracia, ByVal pnInteres As Double, ByVal pnNroCuotas As Integer) As Variant
Dim nValorProrateado As Double
Dim i As Integer
Dim MatGracia() As Double
Dim sMatGracia() As String

    nValorProrateado = pnInteres / pnNroCuotas
    nValorProrateado = CDbl(Format(nValorProrateado, "#0.00"))
    ReDim MatGracia(pnNroCuotas)
    If pOptGracia = PrimeraCuota Then
        ReDim Preserve MatGracia(pnNroCuotas + 1)
        MatGracia(0) = pnInteres
    End If
    If pOptGracia = UltimaCuota Then
        ReDim Preserve MatGracia(pnNroCuotas + 1)
        MatGracia(pnNroCuotas) = pnInteres
    End If
    If pOptGracia = Prorateada Then
        For i = 0 To pnNroCuotas - 1
            If i = pnNroCuotas - 1 Then
                MatGracia(i) = pnInteres - CDbl(Format((nValorProrateado * (pnNroCuotas - 1)), "#0.00"))
            Else
                MatGracia(i) = nValorProrateado
            End If
        Next i
    End If
    
    If pOptGracia = Exonerada Then
        ReDim MatGracia(pnNroCuotas)
    End If
    
    ReDim sMatGracia(UBound(MatGracia))
    For i = 0 To UBound(MatGracia) - 1
        sMatGracia(i) = Format(MatGracia(i), "#0.00")
    Next i
    GeneraGracia = sMatGracia
End Function


Private Sub CalendarioCreciente(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant)


Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As NCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nSumCuotas As Double

    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    nSumCuotas = 0
    
    For i = 1 To pnNroCuotas
        nSumCuotas = nSumCuotas + i
    Next i
    
    If pnTipoPeriodo = PeriodoFijo Then
        Set oCredito = New NCredito
        For i = 0 To pnNroCuotas - 1
            Calendario(i).dFecha = dDesembolso + pnPeriodo
            Calendario(i).NroCuota = i + 1
            Calendario(i).IntGra = 0#
            Calendario(i).Gasto = 0#
            Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            If i = pnNroCuotas - 1 Then
                Calendario(i).Captital = nSaldoCapital
            Else
                Calendario(i).Captital = CDbl(Format((Calendario(i).NroCuota / nSumCuotas) * pnMonto, "#0.00"))
            End If
            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
            Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
            nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            dDesembolso = dDesembolso + pnPeriodo
        Next i
        
        
    Else
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New NCredito
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            nDia = pnDiaFijo
            For i = 0 To pnNroCuotas - 1
                nDia = pnDiaFijo
                If Not (i = 0 And pnDiaFijo > Day(dDesembolso) And Not bProxMes) Then
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
                Calendario(i).dFecha = dFecTemp
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0
                
                If i = 0 Then
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                Else
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha))
                End If
                
                If i = pnNroCuotas - 1 Then
                    Calendario(i).Captital = nSaldoCapital
                Else
                    Calendario(i).Captital = CDbl(Format((Calendario(i).NroCuota / nSumCuotas) * pnMonto, "#0.00"))
                End If
                Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            Next i
        End If
    End If
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For i = pnNroCuotas To 1 Step -1
                Calendario(i) = Calendario(i - 1)
                Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            Next i
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
            For i = 0 To pnNroCuotas - 1
                Calendario(i).IntGra = MatGracia(i)
                Calendario(i).Cuota = Calendario(i).Cuota + MatGracia(i)
            Next i
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
            Calendario(pnNroCuotas).Gasto = 0
            Calendario(pnNroCuotas).IntComp = 0#
            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
            Calendario(pnNroCuotas).Captital = 0#
            Calendario(pnNroCuotas).SaldoCap = 0#
        End If
    End If

End Sub

Private Sub CalendarioDecreciente(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, Optional ByVal MatGracia As Variant)

Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As NCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer

    nSaldoCapital = pnMonto
    dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
    
            
    If pnTipoPeriodo = PeriodoFijo Then
        Set oCredito = New NCredito
        For i = 0 To pnNroCuotas - 1
            Calendario(i).dFecha = dDesembolso + pnPeriodo
            Calendario(i).NroCuota = i + 1
            Calendario(i).IntGra = 0#
            Calendario(i).Gasto = 0#
            Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
            
            If i = pnNroCuotas - 1 Then
                Calendario(i).Captital = nSaldoCapital
            Else
                Calendario(i).Captital = CDbl(Format((pnMonto / pnNroCuotas), "#0.00"))
            End If
            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
            Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
            nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            dDesembolso = dDesembolso + pnPeriodo
        Next i
    Else
        'Si es Fecha Fija
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New NCredito
            nMes = Month(dDesembolso)
            nAnio = Year(dDesembolso)
            nDia = pnDiaFijo
            For i = 0 To pnNroCuotas - 1
                If Not (i = 0 And pnDiaFijo > Day(dDesembolso) And Not bProxMes) Then
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
                Calendario(i).dFecha = dFecTemp
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0
                If i = 0 Then
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                Else
                    Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                    'Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNrocuotas, pnMonto, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha))
                End If
                If i = pnNroCuotas - 1 Then
                    Calendario(i).Captital = nSaldoCapital
                Else
                    Calendario(i).Captital = CDbl(Format((pnMonto / pnNroCuotas), "#0.00"))
                End If
                Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = nSaldoCapital - Calendario(i).Captital
            Next i
        End If
    End If
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For i = pnNroCuotas To 1 Step -1
                Calendario(i) = Calendario(i - 1)
                Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            Next i
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
            For i = 0 To pnNroCuotas - 1
                Calendario(i).IntGra = MatGracia(i)
                Calendario(i).Cuota = Calendario(i).Cuota + MatGracia(i)
            Next i
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
            Calendario(pnNroCuotas).Gasto = 0
            Calendario(pnNroCuotas).IntComp = 0#
            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
            Calendario(pnNroCuotas).Captital = 0#
            Calendario(pnNroCuotas).SaldoCap = 0#
        End If
    End If
End Sub
'
'Private Sub CalendarioCuotaFija(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
'                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
'                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
'                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
'                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
'                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False)
'
'Dim nSaldoCapital As Double
'Dim dDesembolso As Date
'Dim I As Integer
'Dim oCredito As NCredito
'Dim dFecTemp As Date
'Dim nMes As Integer
'Dim nAnio As Integer
'Dim nDia As Integer
'Dim nMontoCuotaTemp As Double
'Dim nCDIntPend As Double
'Dim nPeriodoDesPar As Integer
'Dim oCalend As Dcalendario
'
'    If pbDesemParcial Then
'        Set oCredito = New NCredito
'        ReDim Calendario(1)
'        Set oCalend = New Dcalendario
'        Calendario(0).dFecha = CDate(pMatDesPar(UBound(pMatDesPar) - 1, 2)) + pnPeriodo
'        Set oCalend = Nothing
'        Calendario(0).IntComp = 0
'        Calendario(0).Captital = 0
'        nSaldoCapital = 0
'        For I = 0 To UBound(pMatDesPar) - 1
'            nSaldoCapital = CDbl(pMatDesPar(I, 1))
'            Calendario(0).Captital = Calendario(0).Captital + nSaldoCapital
'            nPeriodoDesPar = DateDiff("d", CDate(pMatDesPar(I, 0)), Calendario(0).dFecha)
'            Calendario(0).IntComp = Calendario(0).IntComp + oCredito.MontoIntPerDias(pnTasaInt, nPeriodoDesPar, nSaldoCapital)
'        Next I
'        Calendario(0).NroCuota = 1
'
'        Calendario(0).IntGra = 0#
'        Calendario(0).Gasto = 0#
'        Calendario(0).Cuota = CDbl(Format(pnMonto + Calendario(0).IntComp, "#0.00"))
'        Set oCredito = Nothing
'    Else
'
'        nSaldoCapital = pnMonto
'        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
'        If pnTipoPeriodo = PeriodoFijo Then
'            Set oCredito = New NCredito
'            For I = 0 To pnNroCuotas - 1
'                Calendario(I).dFecha = dDesembolso + pnPeriodo
'                Calendario(I).NroCuota = I + 1
'                Calendario(I).IntGra = 0#
'                Calendario(I).Gasto = 0#
'                Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
'                Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo)
'                If I = pnNroCuotas - 1 Then
'                    Calendario(I).Captital = nSaldoCapital
'                    If (Calendario(I).IntComp + Calendario(I).Captital) > Calendario(I).Cuota Then
'                        Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
'                    Else
'                        Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
'                    End If
'                Else
'                    Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
'                End If
'                Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
'                nSaldoCapital = nSaldoCapital - Calendario(I).Captital
'                nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
'                dDesembolso = dDesembolso + pnPeriodo
'            Next I
'
'        Else
'            If pbCuotaFijaFechaFija Then
'                Set oCredito = New NCredito
'                nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv)
'                Set oCredito = Nothing
'            End If
'            'Si es Fecha Fija
'            If pnTipoPeriodo = FechaFija Then
'                Set oCredito = New NCredito
'                nMes = Month(dDesembolso)
'                nAnio = Year(dDesembolso)
'                nDia = pnDiaFijo
'                nCDIntPend = 0
'                For I = 0 To pnNroCuotas - 1
'                    nDia = pnDiaFijo
'                    If Not (I = 0 And pnDiaFijo > Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then
'                        nMes = nMes + pnNumMes
'                        If nMes > 12 Then
'                            nAnio = nAnio + 1
'                            nMes = nMes - 12
'                        End If
'                    Else
'                        If nMes = 2 Then
'                            If nDia >= 29 Then
'                                If nAnio Mod 4 <> 0 Then
'                                    nMes = nMes + pnNumMes
'                                End If
'                            End If
'                        Else
'                            If nDia > 30 Then
'                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
'                                    nMes = nMes + pnNumMes
'                                End If
'                            End If
'                        End If
'                    End If
'                    If nMes = 2 Then
'                        If nDia > 28 Then
'                            If nAnio Mod 4 = 0 Then
'                                nDia = 29
'                            Else
'                                nDia = 28
'                            End If
'                        End If
'                    Else
'                        If nDia > 30 Then
'                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or 11 Then
'                                nDia = 30
'                            End If
'                        End If
'                    End If
'                    dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
'                    Calendario(I).dFecha = dFecTemp
'                    Calendario(I).NroCuota = I + 1
'                    Calendario(I).IntGra = 0#
'                    Calendario(I).Gasto = 0
'                    If I = 0 Then
'                        If pbCuotaFijaFechaFija And pbMiViv Then
'                            dDesembolso = dDesembolso - pnDiasGracia
'                        End If
'                        Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(I).dFecha), nSaldoCapital)
'                        If Not pbCuotaFijaFechaFija Then
'                            Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(I).dFecha))
'                        Else
'                            Calendario(I).Cuota = nMontoCuotaTemp
'                        End If
'
'                    Else
'                        Calendario(I).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha), nSaldoCapital)
'                        If Not pbCuotaFijaFechaFija Then
'                            Calendario(I).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", Calendario(I - 1).dFecha, Calendario(I).dFecha))
'                        Else
'                            Calendario(I).Cuota = nMontoCuotaTemp
'                        End If
'                    End If
'
'                    If pbCuotaFijaFechaFija Then
'                        Calendario(I).IntComp = Calendario(I).IntComp + nCDIntPend
'                        nCDIntPend = 0#
'                        If Calendario(I).IntComp > nMontoCuotaTemp Then
'                            nCDIntPend = Calendario(I).IntComp - nMontoCuotaTemp
'                            Calendario(I).IntComp = nMontoCuotaTemp
'                        End If
'                    End If
'
'                    If I = pnNroCuotas - 1 Then
'                        Calendario(I).Captital = nSaldoCapital
'                        If (Calendario(I).IntComp + Calendario(I).Captital) > Calendario(I).Cuota Then
'                            Calendario(I).Cuota = Calendario(I).Captital + Calendario(I).IntComp
'                        Else
'                            Calendario(I).IntComp = Calendario(I).Cuota - Calendario(I).Captital
'                        End If
'                    Else
'                        Calendario(I).Captital = Calendario(I).Cuota - Calendario(I).IntComp
'                    End If
'
'                    Calendario(I).SaldoCap = nSaldoCapital - Calendario(I).Captital
'                    nSaldoCapital = nSaldoCapital - Calendario(I).Captital
'                Next I
'            End If
'        End If
'    End If
'
'    'Actualizar si existe Periodo de Gracia
'    If pnDiasGracia > 0 Then
'        If pnTipoGracia = PrimeraCuota Then
'            ReDim Preserve Calendario(pnNroCuotas + 1)
'            For I = pnNroCuotas To 1 Step -1
'                Calendario(I) = Calendario(I - 1)
'                Calendario(I).NroCuota = Calendario(I).NroCuota + 1
'            Next I
'            Calendario(0).dFecha = pdFecDesemb + pnDiasGracia
'            Calendario(0).NroCuota = 1
'            Calendario(0).IntGra = MatGracia(0)
'            Calendario(0).Gasto = 0
'            Calendario(0).IntComp = 0#
'            Calendario(0).Cuota = Calendario(0).IntGra
'            Calendario(0).Captital = 0#
'            Calendario(0).SaldoCap = pnMonto
'        End If
'        If pnTipoGracia = Prorateada Or pnTipoGracia = Configurable Then
'            For I = 0 To pnNroCuotas - 1
'                Calendario(I).IntGra = MatGracia(I)
'                Calendario(I).Cuota = Calendario(I).Cuota + MatGracia(I)
'            Next I
'        End If
'        If pnTipoGracia = UltimaCuota Then
'            ReDim Preserve Calendario(pnNroCuotas + 1)
'            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
'            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
'            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
'            Calendario(pnNroCuotas).Gasto = 0
'            Calendario(pnNroCuotas).IntComp = 0#
'            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
'            Calendario(pnNroCuotas).Captital = 0#
'            Calendario(pnNroCuotas).SaldoCap = 0#
'        End If
'    End If
'End Sub


Private Sub CalendarioCuotaFija(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False)
                
Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As NCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double
Dim nCDIntPend As Double
Dim nPeriodoDesPar As Integer

    If pbDesemParcial Then
        Set oCredito = New NCredito
        ReDim Calendario(1)
        
        'Calendario(0).dFecha = CDate(pMatDesPar(UBound(pMatDesPar) - 1, 0)) + pnPeriodo
        
        'Cambiado para Santa -- LAYG
        'Valida que la fecha sea menor que el ultimo desemboloso
        'If CDate(pdFecDesemb) + pnPeriodo > CDate(pMatDesPar(UBound(pMatDesPar) - 1, 0)) Then
        '    MsgBox "Fecha de Ultimo desembolso es Posterior al Plazo del credito", vbInformation, "Aviso"
        'End If
        Calendario(0).dFecha = CDate(pdFecDesemb) + pnPeriodo
         
        Calendario(0).IntComp = 0
        Calendario(0).Captital = 0
        nSaldoCapital = 0
        For i = 0 To UBound(pMatDesPar) - 1
            nSaldoCapital = CDbl(pMatDesPar(i, 1))
            Calendario(0).Captital = Calendario(0).Captital + nSaldoCapital
            nPeriodoDesPar = DateDiff("d", CDate(pMatDesPar(i, 0)), Calendario(0).dFecha)
            Calendario(0).IntComp = Calendario(0).IntComp + oCredito.MontoIntPerDias(pnTasaInt, nPeriodoDesPar, nSaldoCapital)
        Next i
        Calendario(0).NroCuota = 1
        
        Calendario(0).IntGra = 0#
        Calendario(0).Gasto = 0#
        Calendario(0).Cuota = CDbl(Format(pnMonto + Calendario(0).IntComp, "#0.00"))
        Set oCredito = Nothing
    Else

        nSaldoCapital = pnMonto
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
        If pnTipoPeriodo = PeriodoFijo Then
            Set oCredito = New NCredito
            For i = 0 To pnNroCuotas - 1
                Calendario(i).dFecha = dDesembolso + pnPeriodo
                Calendario(i).NroCuota = i + 1
                Calendario(i).IntGra = 0#
                Calendario(i).Gasto = 0#
                Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, pnPeriodo, nSaldoCapital)
                Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo)
                If i = pnNroCuotas - 1 Then
                    Calendario(i).Captital = nSaldoCapital
                    If (Calendario(i).IntComp + Calendario(i).Captital) > Calendario(i).Cuota Then
                        Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                    Else
                        Calendario(i).IntComp = Calendario(i).Cuota - Calendario(i).Captital
                    End If
                Else
                    Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp
                End If
                Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = nSaldoCapital - Calendario(i).Captital
                nSaldoCapital = CDbl(Format(nSaldoCapital, "#0.00"))
                dDesembolso = dDesembolso + pnPeriodo
            Next i
            
        Else
            If pbCuotaFijaFechaFija Then
                Set oCredito = New NCredito
                nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, pnDiaFijo, pnDiasGracia, bProxMes, pnNumMes, pbMiViv)
                Set oCredito = Nothing
            End If
            'Si es Fecha Fija
            If pnTipoPeriodo = FechaFija Then
                Set oCredito = New NCredito
                nMes = Month(dDesembolso)
                nAnio = Year(dDesembolso)
                nDia = pnDiaFijo
                nCDIntPend = 0
                For i = 0 To pnNroCuotas - 1
                    nDia = pnDiaFijo
                    If Not (i = 0 And pnDiaFijo > Day(dDesembolso) And (Not bProxMes)) Or pnNumMes = 6 Then
                        nMes = nMes + pnNumMes
                        If nMes > 12 Then
                            nAnio = nAnio + 1
                            nMes = 1
                        End If
                    Else
                        If nMes = 2 Then
                            If nDia >= 29 Then
                                If nAnio Mod 4 <> 0 Then
                                    nMes = nMes + pnNumMes
                                End If
                            End If
                        Else
                            If nDia > 30 Then
                                If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                                    nMes = nMes + pnNumMes
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
                            If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
                                nDia = 30
                            End If
                        End If
                    End If
                    dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes)), 2) & "/" & Trim(Str(nAnio)))
                    Calendario(i).dFecha = dFecTemp
                    Calendario(i).NroCuota = i + 1
                    Calendario(i).IntGra = 0#
                    Calendario(i).Gasto = 0
                    If i = 0 Then
                        If pbCuotaFijaFechaFija And pbMiViv Then
                            dDesembolso = dDesembolso - pnDiasGracia
                        End If
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", dDesembolso, Calendario(i).dFecha), nSaldoCapital)
                        If Not pbCuotaFijaFechaFija Then
                            Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", dDesembolso, Calendario(i).dFecha))
                        Else
                            Calendario(i).Cuota = nMontoCuotaTemp
                        End If
                        
                    Else
                        Calendario(i).IntComp = oCredito.MontoIntPerDias(pnTasaInt, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha), nSaldoCapital)
                        If Not pbCuotaFijaFechaFija Then
                            Calendario(i).Cuota = oCredito.CuotaFija(pnTasaInt, pnNroCuotas, pnMonto, DateDiff("d", Calendario(i - 1).dFecha, Calendario(i).dFecha))
                        Else
                            Calendario(i).Cuota = nMontoCuotaTemp
                        End If
                    End If
                    
                    If pbCuotaFijaFechaFija Then
                        Calendario(i).IntComp = Calendario(i).IntComp + nCDIntPend
                        nCDIntPend = 0#
                        If Calendario(i).IntComp > nMontoCuotaTemp Then
                            nCDIntPend = Calendario(i).IntComp - nMontoCuotaTemp
                            Calendario(i).IntComp = nMontoCuotaTemp
                        End If
                    End If
                    
                    If i = pnNroCuotas - 1 Then
                        Calendario(i).Captital = nSaldoCapital
                        If (Calendario(i).IntComp + Calendario(i).Captital) > Calendario(i).Cuota Then
                            Calendario(i).Cuota = Calendario(i).Captital + Calendario(i).IntComp
                        Else
                            Calendario(i).IntComp = Calendario(i).Cuota - Calendario(i).Captital
                        End If
                    Else
                        Calendario(i).Captital = Calendario(i).Cuota - Calendario(i).IntComp
                    End If
                    
                    Calendario(i).SaldoCap = nSaldoCapital - Calendario(i).Captital
                    nSaldoCapital = nSaldoCapital - Calendario(i).Captital
                Next i
            End If
        End If
    End If
    
    'Actualizar si existe Periodo de Gracia
    If pnDiasGracia > 0 Then
        If pnTipoGracia = PrimeraCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            For i = pnNroCuotas To 1 Step -1
                Calendario(i) = Calendario(i - 1)
                Calendario(i).NroCuota = Calendario(i).NroCuota + 1
            Next i
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
            For i = 0 To pnNroCuotas - 1
                Calendario(i).IntGra = MatGracia(i)
                Calendario(i).Cuota = Calendario(i).Cuota + MatGracia(i)
            Next i
        End If
        If pnTipoGracia = UltimaCuota Then
            ReDim Preserve Calendario(pnNroCuotas + 1)
            Calendario(pnNroCuotas).dFecha = Calendario(pnNroCuotas - 1).dFecha + pnDiasGracia
            Calendario(pnNroCuotas).NroCuota = pnNroCuotas + 1
            Calendario(pnNroCuotas).IntGra = MatGracia(pnNroCuotas)
            Calendario(pnNroCuotas).Gasto = 0
            Calendario(pnNroCuotas).IntComp = 0#
            Calendario(pnNroCuotas).Cuota = Calendario(pnNroCuotas).IntGra
            Calendario(pnNroCuotas).Captital = 0#
            Calendario(pnNroCuotas).SaldoCap = 0#
        End If
    End If
End Sub
Private Sub CalendarioQuincena(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False)
                
Dim nSaldoCapital As Double
Dim dDesembolso As Date
Dim i As Integer
Dim oCredito As NCredito
Dim dFecTemp As Date
Dim nMes As Integer
Dim nAnio As Integer
Dim nDia As Integer
Dim nMontoCuotaTemp As Double
Dim nCDIntPend As Double
Dim nPeriodoDesPar As Integer
'Considerando que tiene 2 fechas el 16 y el 1 de cada mes
Dim nDia1 As Integer
Dim nDia2 As Integer
Dim bCambioFechas As Boolean

        nDia1 = 1
        nDia2 = 16
        nSaldoCapital = nMonto
        dDesembolso = CDate(Format(pdFecDesemb, "dd/mm/yyyy")) + pnDiasGracia
        
        
        Set oCredito = New NCredito
        nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, nDia1, pnDiasGracia, bProxMes, pnNumMes, pbMiViv)
        nMontoCuotaTemp = oCredito.CFijaFechaFija(pnTasaInt, pnNroCuotas, pnMonto, pnPeriodo, pdFecDesemb, nDia2, pnDiasGracia, bProxMes, pnNumMes, pbMiViv)
        Set oCredito = Nothing
        
        If pnTipoPeriodo = FechaFija Then
            Set oCredito = New NCredito
                nMes = Month(dDesembolso)
                nAnio = Year(dDesembolso)
                nDia = pnDiaFijo
                nCDIntPend = 0
                
                bCambioFechas = False
            
                For i = 0 To pnNroCuotas - 1
                    
                    If bCambioFechas = False Then
                        ' el menor dia
                        bCambioFechas = True
                    Else
                        ' el mayor dia
                        bCambioFechas = False
                    End If
                Next i
        End If
        
End Sub


Private Sub ProcesarCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, Optional ByVal bQuincena As Boolean)
                
                Select Case pnTipoCuota
                    Case Creciente
                        Call CalendarioCreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia)
                    Case Decreciente
                        Call CalendarioDecreciente(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia)
                    Case Fija
                        Call CalendarioCuotaFija(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, pMatDesPar, pnNumMes, pbMiViv)
                End Select
                
                If bQuincena = True Then
                
                End If
End Sub
Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis

End Sub
Public Function GeneraCalendario(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Double, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnDiaFijo As Integer, ByVal bProxMes As Boolean, _
                Optional ByVal MatGracia As Variant, Optional ByVal pbCuotaFijaFechaFija As Boolean = False, _
                Optional ByVal pbCuotaComodin As Boolean = False, Optional ByVal pbDesemParcial As Boolean = False, _
                Optional ByVal pMatDesPar As Variant = "", Optional ByVal pnNumMes As Integer = 1, Optional ByVal pbMiViv As Boolean = False, _
                Optional ByVal bQuincena As Boolean) As Variant
Dim sCalendPagos() As String
Dim i As Integer
Dim TasaComTemp As Double
Dim nPlazoCom As Integer
Dim nCuotasCom As Integer
Dim nCuotaMontoCom As Double
Dim oCred As NCredito

        On Error GoTo ErrorGeneraCalendario
        

        If pbCuotaComodin Then
            TasaComTemp = pnTasaInt
            If pnTipoPeriodo = PeriodoFijo Then
                nPlazoCom = pnPeriodo
            Else
                nPlazoCom = 30
            End If
            nCuotasCom = pnNroCuotas
            Set oCred = New NCredito
            nCuotaMontoCom = oCred.CuotaFija(TasaComTemp, nCuotasCom + 1, pnMonto, nPlazoCom)
            nCuotaMontoCom = CDbl(Format((nCuotaMontoCom * (pnNroCuotas + 1)) / pnNroCuotas, "#0.00"))
            TasaComTemp = CDbl(Format(oCred.CreditosTasaEfectiva(nCuotaMontoCom, pnNroCuotas, pnMonto, pnTasaInt, nPlazoCom), "#0.0000"))
            pnTasaInt = Format(TasaComTemp * 100, "#0.0000")
            Set oCred = Nothing
        End If

        
        ReDim Calendario(pnNroCuotas)
        Call ProcesarCalendario(pnMonto, pnTasaInt, pnNroCuotas, pnPeriodo, pdFecDesemb, pnTipoCuota, _
                pnTipoPeriodo, pnTipoGracia, pnDiasGracia, pnDiaFijo, bProxMes, MatGracia, pbCuotaFijaFechaFija, pbDesemParcial, pMatDesPar, pnNumMes, pbMiViv)
        ReDim sCalendPagos(UBound(Calendario), 8)
        For i = 0 To UBound(Calendario) - 1
            sCalendPagos(i, 0) = Format(Calendario(i).dFecha, "dd/mm/yyyy")
            sCalendPagos(i, 1) = Trim(Str(Calendario(i).NroCuota))
            sCalendPagos(i, 2) = Format(Calendario(i).Cuota, "#0.00")
            sCalendPagos(i, 3) = Format(Calendario(i).Captital, "#0.00")
            sCalendPagos(i, 4) = Format(Calendario(i).IntComp, "#0.00")
            sCalendPagos(i, 5) = Format(Calendario(i).IntGra, "#0.00")
            sCalendPagos(i, 6) = Format(Calendario(i).Gasto, "#0.00")
            sCalendPagos(i, 7) = Format(Calendario(i).SaldoCap, "#0.00")
        Next i
        GeneraCalendario = sCalendPagos
        Exit Function

ErrorGeneraCalendario:
        MsgBox Err.Description, vbCritical, "Aviso"
        
End Function

Public Function ReporteCalendario(ByVal pTipoRep As Byte, ByVal MatP As Variant, ByVal MatMalo As Variant, _
        ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date, _
        Optional ByVal pnTipoProceso As Integer = 0, Optional ByVal pMatDesPar As Variant = "", Optional GBITFAPLICAt As Boolean, Optional gnITFPorcentt As Double, Optional gnITFMontoMint As Double = 0) As String

Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim nItem As Integer

Dim cTitulo As String
Dim i As Integer
Dim cCuota As String * 35
Dim sCapital As Double
Dim sInteres As Double
 
gbITFAplica = GBITFAPLICAt
gnITFPorcent = gnITFPorcentt
gnITFMontoMin = gnITFMontoMint
 
cCuota = pTCuota
sCapital = 0
sInteres = 0
  
    If pTipoRep = 1 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION DE CALENDARIO DE PAGOS"
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS"
        End If
    ElseIf pTipoRep = 2 Then
        If pnTipoProceso = 0 Then
            cTitulo = "SIMULACION DE CALENDARIO DE PAGOS MI VIVIENDA"
        ElseIf pnTipoProceso = 1 Then
            cTitulo = "SUGERENCIA DE ANALISTA - CALENDARIO DE PAGOS MI VIVIENDA"
        ElseIf pnTipoProceso = 2 Then
            cTitulo = "APROBACION DE CREDITO - CALENDARIO DE PAGOS MI VIVIENDA"
        End If
    End If

    lnLineas = 0
    lnPage = 1

    'CABECERA
    If pTipoRep = 1 Then
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "PLAN DE PAGOS", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
    ElseIf pTipoRep = 2 Then
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "BUEN PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
    End If
    
    
    If IsArray(pMatDesPar) Then
        lsCadImp = lsCadImp & Space(5) & "DESEMBOLSOS PARCIALES" & Chr(10)
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(5) & ImpreFormat("FECHA", 16, 0) & ImpreFormat("MONTO", 10, 0) & Chr(10)
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & Chr(10)
        For i = 0 To UBound(pMatDesPar) - 1
            lsCadImp = lsCadImp & Space(5) & ImpreFormat(pMatDesPar(i, 0), 10, 0) & ImpreFormat(CDbl(pMatDesPar(i, 1)), 10, 2, True) & Chr(10)
        Next i
        lsCadImp = lsCadImp & Space(5) & String(63, "-") & Chr(10)
    End If
    
    lsCadImp = lsCadImp & String(123, "-") & Chr(10)
    lsCadImp = lsCadImp & "ITEM  FECHA            NO.CUOTA        CUOTA      CAPITAL      INTERES   INT.GRACIA       GASTOS   SALDO CAP.   CUOTA + ITF" & Chr(10)
    lsCadImp = lsCadImp & String(123, "-") & Chr(10)
    
    'lsCadImp = lsCadImp & Chr(10)
    
    lnIndice = 10:  lnLineas = 10
    nItem = 0
    For i = 0 To UBound(MatP) - 1
    
        nItem = nItem + 1
        lnLineas = lnLineas + 1
        lnIndice = lnIndice + 1
        lsCadImp = lsCadImp & ImpreFormat(nItem, 4, 0) & Space(2)
        lsCadImp = lsCadImp & ImpreFormat(Format(MatP(i, 0), "ddd, dd mmm yyyy"), 16, 0) & Space(2)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 1)), 7, 0) & Space(2)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 2)), 8, 2) & Space(2)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 3)), 8, 2) & Space(2)
        sCapital = sCapital + MatP(i, 3)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 4)), 8, 2) & Space(2)
        sInteres = sInteres + MatP(i, 4)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 5)), 8, 2) & Space(2)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 6)), 8, 2) & Space(2)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 7)), 8, 2) & Space(2)
        lsCadImp = lsCadImp & ImpreFormat(Val(MatP(i, 2)) + fgITFCalculaImpuesto(Val(MatP(i, 2))), 8, 2) & Space(2) & Chr(10)
 
        If lnIndice Mod 300 = 0 Then
            lsCadBuffer = lsCadBuffer & lsCadImp
            lsCadImp = ""
        End If
        
        If lnLineas >= 55 Then
            lnPage = lnPage + 1
            lsCadImp = lsCadImp & Chr(12)
            lsCadImp = lsCadImp & Chr(10)
            If pTipoRep = 1 Then
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "PLAN DE PAGOS", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
            ElseIf pTipoRep = 2 Then
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "BUEN PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
            End If
        
            lnLineas = 8
            lnIndice = lnIndice + 8
        End If
                
    Next

    lsCadImp = lsCadImp & Chr(10) & String(123, "=") & Chr(10)
    lsCadImp = lsCadImp & "Totales :" & Space(37) & ImpreFormat(sCapital, 8, 2)
    lsCadImp = lsCadImp & Space(2) & ImpreFormat(sInteres, 8, 2)
    lsCadImp = lsCadImp & Space(28) & ImpreFormat(sCapital + sInteres, 8, 2)

    'Mal Pagador Mi Vivienda
    
    sCapital = 0
    sInteres = 0
     
    If pTipoRep = 2 Then
        lsCadImp = lsCadImp & Chr(12)
        lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
     
        lsCadImp = lsCadImp & Chr(10)
        
        lnIndice = 10:  lnLineas = 10
        nItem = 0
        For i = 0 To UBound(MatMalo) - 1
            nItem = nItem + 1
            lnLineas = lnLineas + 1
            lnIndice = lnIndice + 1
            lsCadImp = lsCadImp & ImpreFormat(nItem, 4, 0) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(Format(MatMalo(i, 0), "ddd, dd mmm yyyy"), 16, 0) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 1)), 7, 0) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 2)), 8, 2) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 3)), 8, 2) & Space(2)
            sCapital = sCapital + MatMalo(i, 3)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 4)), 8, 2) & Space(2)
            sInteres = sInteres + MatMalo(i, 4)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 5)), 8, 2) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 6)), 8, 2) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 7)), 8, 2) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(Val(MatMalo(i, 8)), 8, 2) & Space(2) & Chr(10)
              
            'lnLineas = lnLineas + 1
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & lsCadImp
                lsCadImp = ""
            End If
            
            If lnLineas >= 55 Then
                lnPage = lnPage + 1
                lsCadImp = lsCadImp & Chr(12)
                lsCadImp = lsCadImp & Chr(10)
                lsCadImp = lsCadImp & nRepoCabecera(cTitulo, "MAL PAGADOR", lnPage, 109, "", "CalenPagos", cCuota, pInteres, pMonto, pCuota, pPlazo, pVigencia)
             
                lnLineas = 8
                lnIndice = lnIndice + 8
            End If
                    
        Next
        
        lsCadImp = lsCadImp & Chr(10) & String(109, "=") & Chr(10)
        lsCadImp = lsCadImp & "Totales :" & Space(37) & ImpreFormat(sCapital, 8, 2)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat(sInteres, 8, 2)
        lsCadImp = lsCadImp & Space(28) & ImpreFormat(sCapital + sInteres, 8, 2)

    End If
ReporteCalendario = lsCadBuffer & lsCadImp
 
 End Function
Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As String, ByVal pTCuota As String, ByVal pInteres As Double, ByVal pMonto As Double, _
        ByVal pCuota As Double, ByVal pPlazo As Integer, ByVal pVigencia As Date) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NColPImpre

    Set loImprimeCab = New NColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    Select Case pnCodReporte
        Case "CalenPagos"
            lsCadImp = lsCadImp & "    TIPO DE CUOTA   : " & pTCuota & Space(20) & "CUOTA :      " & ImpreFormat(pMonto, 10, 2) & Chr(10)
            lsCadImp = lsCadImp & "    INTERES         : " & ImpreFormat(pInteres, 8, 4) & Space(42) & "PLAZO :      " & ImpreFormat(pPlazo, 10, 2) & Chr(10)
            lsCadImp = lsCadImp & "    MONTO           :   " & ImpreFormat(pMonto, 8, 2) & Space(42) & "VIGENCIA: " & Format(pVigencia, "ddd, dd mmm yyyy") & Chr(10)
            lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
            
       End Select
       
nRepoCabecera = lsCadImp
End Function


 
